from decimal import Decimal

import pandas as pd

from data_loader import DataLoader, ParseErrorCollector


def _标准化(借方, 贷方):
    原表 = pd.DataFrame(
        [{"日期": "2026-01-02", "摘要": "金额缺失", "借方": 借方, "贷方": 贷方}]
    )
    收集器 = ParseErrorCollector()
    结果 = DataLoader(error_collector=收集器).standardize_data(
        原表,
        {
            "date": "日期",
            "summary": "摘要",
            "mode": "debit_credit",
            "debit": "借方",
            "credit": "贷方",
        },
        "bank",
    )
    return 结果, 收集器.get_all_errors()


def test_借贷两格都空必须作为金额缺失排除():
    原表 = pd.DataFrame(
        [
            {"日期": "2026-01-02", "摘要": "金额缺失", "借方": "", "贷方": ""},
            {"日期": "2026-01-03", "摘要": "正常交易", "借方": "", "贷方": 100},
        ]
    )
    收集器 = ParseErrorCollector()
    结果 = DataLoader(error_collector=收集器).standardize_data(
        原表,
        {
            "date": "日期",
            "summary": "摘要",
            "mode": "debit_credit",
            "debit": "借方",
            "credit": "贷方",
        },
        "bank",
    )
    错误 = 收集器.get_all_errors()

    assert len(结果) == 1
    assert 结果.iloc[0]["summary"] == "正常交易"
    assert len(错误) == 1
    assert 错误[0]["type"] == "金额解析失败"
    assert "均为空" in str(错误[0]["original_value"])


def test_明确填写零金额仍按零金额交易处理():
    结果, 错误 = _标准化(0, "")

    assert 错误 == []
    assert len(结果) == 1
    assert 结果.iloc[0]["amount"] == Decimal("0.00")
