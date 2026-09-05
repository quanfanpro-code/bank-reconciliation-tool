from decimal import Decimal

import pandas as pd

from data_loader import DataLoader
from data_structures import MatcherConfig
from matcher import Matcher
from reporter import Reporter


def _映射():
    return {
        "date": "交易日期",
        "summary": "摘要",
        "mode": "debit_credit",
        "debit": "本币借方金额",
        "credit": "本币贷方金额",
        "amount_basis": "本位币",
        "account": "账号",
        "currency": "币种",
        "auxiliary_text_columns": ["摘要", "业务编号"],
    }


def _原表(来源, 单边金额):
    if 来源 == "bank":
        借方, 贷方 = [0, 0], [100, 单边金额]
    else:
        借方, 贷方 = [100, 单边金额], [0, 0]
    return pd.DataFrame(
        {
            "交易日期": ["2026-08-31", "2026-01-01" if 来源 == "bank" else "2026-12-31"],
            "入账日期": ["2026-09-01", "2026-01-02" if 来源 == "bank" else "2027-01-01"],
            "起息日": ["2026-09-01", "2026-01-02" if 来源 == "bank" else "2027-01-01"],
            "价值日": ["2026-09-02", "2026-01-03" if 来源 == "bank" else "2027-01-02"],
            "本币借方金额": 借方,
            "本币贷方金额": 贷方,
            "外币借方金额": [0, 0],
            "外币贷方金额": [Decimal("13.88"), Decimal("1.00")],
            "摘要": ["项目回款", "单边项"],
            "业务编号": ["P001", ""],
            "账号": ["A001", "A001"],
            "币种": ["CNY", "CNY"],
        }
    )


def test_标准化结果保留交易入账起息价值日及外币本币原值():
    标准账 = DataLoader().standardize_data(_原表("journal", 60), _映射(), "journal")
    第一行 = 标准账.iloc[0]

    assert 第一行["date_evidence"]["date_column"] == "交易日期"
    assert 第一行["date_evidence"]["related_date_values"] == {
        "交易日期": "2026-08-31",
        "入账日期": "2026-09-01",
        "起息日": "2026-09-01",
        "价值日": "2026-09-02",
    }
    assert 第一行["amount_evidence"]["amount_basis"] == "本位币"
    assert 第一行["amount_evidence"]["related_amount_values"] == {
        "本币借方金额": 100,
        "本币贷方金额": 0,
        "外币借方金额": 0,
        "外币贷方金额": Decimal("13.88"),
    }


def test_匹配组成和双方待查都展开其他业务日期原值():
    加载器 = DataLoader()
    银行 = 加载器.standardize_data(_原表("bank", 50), _映射(), "bank")
    账 = 加载器.standardize_data(_原表("journal", 60), _映射(), "journal")
    匹配器 = Matcher(银行, 账, MatcherConfig(tolerance_days=31))
    匹配器.run()

    表 = Reporter(匹配器).build_report_tables(MatcherConfig(tolerance_days=31))
    for 名称 in ("匹配组成", "银行侧待查", "日记账侧待查"):
        assert not 表[名称].empty
        assert {
            "原日期列名",
            "原日期值",
            "其他业务日期列及原值",
        } <= set(表[名称].columns)
        assert set(表[名称]["原日期列名"]) == {"交易日期"}
        assert 表[名称]["其他业务日期列及原值"].str.contains("入账日期=").all()
        assert 表[名称]["其他业务日期列及原值"].str.contains("起息日=").all()
        assert 表[名称]["其他业务日期列及原值"].str.contains("价值日=").all()
