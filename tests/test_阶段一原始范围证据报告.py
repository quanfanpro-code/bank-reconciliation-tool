from decimal import Decimal

import pandas as pd

from data_loader import DataLoader
from data_structures import MatcherConfig
from matcher import Matcher
from reporter import Reporter


def _原表(来源, 未配金额):
    if 来源 == "bank":
        选中借方, 选中贷方 = [0, 0], [100, 未配金额]
    else:
        选中借方, 选中贷方 = [100, 未配金额], [0, 0]
    return pd.DataFrame({
        "日期": [
            "2026-08-10",
            "2026-01-01" if 来源 == "bank" else "2026-12-31",
        ],
        "本位币借方金额": 选中借方,
        "本位币贷方金额": 选中贷方,
        "原币借方金额": [0, 0],
        "原币贷方金额": [Decimal("13.88"), Decimal("1.00")],
        "摘要": ["项目回款", "年末单边项"],
        "业务编号": ["P001", ""],
        "本方账号": ["A001", "A001"],
        "主币种": ["CNY", "CNY"],
        "原币币种": ["USD", "USD"],
    })


def _映射():
    return {
        "date": "日期",
        "summary": "摘要",
        "mode": "debit_credit",
        "debit": "本位币借方金额",
        "credit": "本位币贷方金额",
        "amount_basis": "本位币",
        "account": "本方账号",
        "currency": "主币种",
        "auxiliary_text_columns": ["摘要", "业务编号"],
    }


def test_匹配组成和双方待查都展开原金额账户及币种证据():
    加载器 = DataLoader()
    银行 = 加载器.standardize_data(_原表("bank", 50), _映射(), "bank")
    账 = 加载器.standardize_data(_原表("journal", 60), _映射(), "journal")
    匹配器 = Matcher(银行, 账, MatcherConfig(tolerance_days=31))
    匹配器.run()

    表 = Reporter(匹配器).build_report_tables(MatcherConfig(tolerance_days=31))
    必需 = {
        "原始相关金额列及原值",
        "原账户列名",
        "原账户值",
        "原币种列名",
        "原币种值",
        "其他币种列及原值",
    }
    for 名称 in ("匹配组成", "银行侧待查", "日记账侧待查"):
        assert not 表[名称].empty
        assert 必需 <= set(表[名称].columns)
        assert set(表[名称]["原账户列名"]) == {"本方账号"}
        assert set(表[名称]["原账户值"]) == {"A001"}
        assert set(表[名称]["原币种列名"]) == {"主币种"}
        assert set(表[名称]["原币种值"]) == {"CNY"}
        assert 表[名称]["原始相关金额列及原值"].str.contains(
            "本位币借方金额="
        ).all()
        assert 表[名称]["原始相关金额列及原值"].str.contains(
            "原币贷方金额="
        ).all()
        assert 表[名称]["其他币种列及原值"].str.contains(
            "原币币种=USD"
        ).all()
