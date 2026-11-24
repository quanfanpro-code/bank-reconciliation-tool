from collections import defaultdict
from decimal import Decimal
from types import SimpleNamespace

import pandas as pd
import pytest

from 核对期间 import apply_coverage, parse_coverage_evidence, uncovered_dates


def views():
    days = pd.Series(pd.to_datetime(["2025-04-28", "2025-04-29", "2025-04-30", "2025-05-01", "2025-05-02"]))
    daily = pd.DataFrame({"日期": days})
    monthly = pd.DataFrame({"月份": days.dt.to_period("M").unique()})
    for prefix in ("银行", "日记账"):
        for key in ("收入金额", "支出金额", "收入笔数", "支出笔数"):
            daily[f"{prefix}-{key}"] = Decimal("0")
            monthly[f"{prefix}-{key}"] = Decimal("0")
        daily[f"{prefix}-期末余额"] = Decimal("120")
        monthly[f"{prefix}-月末余额"] = Decimal("120")
    return days, daily, monthly


@pytest.mark.parametrize("string_months", [False, True])
def test_partial_period_unknown_gross_independent_balance_and_new_anchor(string_months):
    days, daily, monthly = views()
    if string_months:
        monthly["月份"] = monthly["月份"].astype(str)
    data = parse_coverage_evidence({"bank": {"账户乙.xlsx": {
        "balances": [{"date": "2025-04-29", "amount": "96.25", "source": "账户余额第3行"},
                     {"date": "2025-04-30", "amount": "120", "source": "账户余额第4行"}],
        "coverage": [{"start": "2025-04-28", "end": "2025-04-30", "kind": "不完整", "source": "余额文件"},
                     {"start": "2025-04-28", "end": "2025-04-28", "kind": "无交易", "source": "单日查询"}],
    }}})["bank"]["账户乙.xlsx"]
    matcher = SimpleNamespace(bank=pd.DataFrame({"date": days.iloc[-1:]}),
                              journal=pd.DataFrame({"date": days}),
                              uncovered_months={"bank": {pd.Period("2025-04")}}, coverage_evidence={"bank": data})
    daily_notes, monthly_notes = defaultdict(list), defaultdict(list)
    apply_coverage(daily, monthly, matcher, daily_notes, monthly_notes)
    assert daily.at[0, "银行-收入金额"] == 0
    assert pd.isna(daily.at[1, "收入金额差额"])
    assert daily.at[1, "余额差额"] == Decimal("-23.75")
    assert daily.at[2, "余额差额"] == 0
    assert monthly.at[0, "余额差额"] == 0
    assert pd.isna(monthly.at[0, "支出金额差额"])
    assert any("23.75" in note for note in monthly_notes["2025-04"])
    assert any("单日查询" in note for note in daily_notes[days[0]])


def test_opposite_gross_differences_survive_equal_balance():
    days, daily, monthly = views()
    matcher = SimpleNamespace(bank=pd.DataFrame({"date": days}), journal=pd.DataFrame({"date": days}),
                              uncovered_months={})
    for frame in (daily, monthly):
        frame.at[0, "银行-收入金额"] = Decimal("350.25")
        frame.at[0, "银行-支出金额"] = Decimal("350.25")
    daily_notes, monthly_notes = defaultdict(list), defaultdict(list)
    apply_coverage(daily, monthly, matcher, daily_notes, monthly_notes)
    for frame in (daily, monthly):
        assert frame.at[0, "余额差额"] == 0
        assert frame.at[0, "收入金额差额"] == frame.at[0, "支出金额差额"] == Decimal("350.25")
    assert "余额相符，仍有收付差额" in monthly_notes["2025-04"]


def test_undated_or_conflicting_balance_evidence_rejected():
    points = [{"date": "2025-04-07", "amount": "9.25", "source": "第2行"},
              {"date": "2025-04-07", "amount": "9.35", "source": "第3行"}]
    with pytest.raises(ValueError, match="两条不同"):
        parse_coverage_evidence({"bank": {"a.xlsx": {"balances": points}}})
    points[1]["date"] = ""
    with pytest.raises(ValueError):
        parse_coverage_evidence({"bank": {"a.xlsx": {"balances": points}}})


def test_journal_coverage_uses_same_rule_on_a_partial_month():
    days, _, _ = views()
    periods = [(days.iloc[1], days.iloc[2], "不完整", "账上取数范围")]
    assert uncovered_dates(days, (), {"periods": periods}).tolist() == [False, True, True, False, False]

def test_balance_evidence_can_distinguish_same_named_files():
    parsed = parse_coverage_evidence({
        "bank": {"本期.xlsx": {"balances": [{"date": "2025-08-01", "amount": 50, "source": "银行余额表"}]}},
        "journal": {"本期.xlsx": {"balances": [{"date": "2025-08-01", "amount": 45, "source": "账上余额表"}]}},
    })
    day = pd.Timestamp("2025-08-01")
    assert parsed["bank"]["本期.xlsx"]["balances"][day][0] == 50
    assert parsed["journal"]["本期.xlsx"]["balances"][day][0] == 45
