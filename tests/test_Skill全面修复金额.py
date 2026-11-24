from decimal import Decimal

import pandas as pd
import pytest

from data_loader import DataLoader
from data_structures import MatcherConfig
from reporter import Reporter
from 列识别 import recognize_columns
from 运行核对 import SkillMatcher


def 标准化(记录, 来源):
    raw = pd.DataFrame(记录)
    got = recognize_columns(raw, 来源)
    assert not got.choices and not got.problems, got
    return DataLoader().standardize_data(raw, got.mapping, 来源)


@pytest.mark.parametrize("amounts,directions", [([100, -40, -60], None), ([100, 40, 60], ["收款", "付款", "付款"]),
                                              ([100, -40, -60], ["收款", "付款", "付款"])])
def test_同一笔现金变动在不同字段布局下相同(amounts, directions):
    raw = {"日期": ["2025-05-01", "2025-05-02", "2025-05-03"], "交易金额": amounts, "余额": [1100, 1060, 1000]}
    if directions:
        raw["收支方向"] = directions
    result = 标准化(raw, "bank")
    assert result["amount"].tolist() == [Decimal("100"), Decimal("-40"), Decimal("-60")]


def test_原表双向金额不相抵():
    bank = 标准化({"日期": ["2025-05-01"], "收入金额": [160], "支出金额": [60]}, "bank")
    ledger = 标准化({"日期": ["2025-05-01"], "交易金额": [100]}, "journal")
    matcher = SkillMatcher(bank, ledger, MatcherConfig(), logger=lambda _: None)
    result = Reporter(matcher).calculate_daily_stats(bank, ledger).iloc[0]
    assert result["income_amount_diff"] == 60
    assert result["expense_amount_diff"] == 60


def test_金额本身提供的币种随行保留():
    bank = 标准化({"日期": ["2025-05-01"], "交易金额": ["USD 100"]}, "bank")
    ledger = 标准化({"日期": ["2025-05-01"], "交易金额": ["CNY 100"]}, "journal")
    assert bank.at[0, "currency"] == "USD" and ledger.at[0, "currency"] == "CNY"
    matcher = SkillMatcher(bank, ledger, MatcherConfig(), logger=lambda _: None)
    matcher.run()
    assert not matcher.selected_candidates


def test_原金额缺失可用连续余额还原():
    raw = pd.DataFrame({"日期": ["2025-05-01", "2025-05-02", "2025-05-03"],
                        "交易金额": [100, "待确认", 50], "余额": [1100, 1060, 1110]})
    mapping = recognize_columns(raw, "bank").mapping
    result = DataLoader().standardize_data(raw, mapping, "bank", recover_amounts=True)
    assert result["amount"].tolist() == [Decimal("100"), Decimal("-40"), Decimal("50")]
    assert result.at[1, "amount_evidence"]["balance_change"] == Decimal("-40")
