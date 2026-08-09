"""匹配可靠性补强测试。"""

from decimal import Decimal

import pandas as pd

from data_structures import (
    GroupMetrics,
    MatchCandidate,
    MatcherConfig,
    ProcessingStatus,
    TextEvidence,
)
from matcher import Matcher
from matching_policy import route_candidate
from precision_engine import PrecisionEngine


def _exact_candidate(**changes):
    candidate = MatchCandidate(
        candidate_id="exact",
        bank_idxs=(0,),
        journal_idxs=(0,),
        match_type="exact_1to1",
        match_stage="精确",
        metrics=GroupMetrics(
            bank_gross_li=1_000_000,
            journal_gross_li=1_000_000,
            group_amount_li=1_000_000,
            bank_income_li=1_000_000,
            journal_income_li=1_000_000,
            bank_expense_li=0,
            journal_expense_li=0,
            income_diff_li=0,
            expense_diff_li=0,
            total_diff_li=0,
        ),
    )
    for name, value in changes.items():
        setattr(candidate, name, value)
    return candidate


def test_无风险且金额完全一致仍可自动确认():
    candidate = _exact_candidate()

    status, _ = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.AUTO_CONFIRMED


def test_关键文字字段冲突时金额一致也必须复核():
    candidate = _exact_candidate(
        text_evidence=TextEvidence(conflicting_fields=("对方户名",)),
    )

    status, reason = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.PENDING_REVIEW
    assert "对方户名" in reason


def test_候选歧义和文字冲突同时写入复核原因():
    candidate = _exact_candidate(
        is_ambiguous=True,
        text_evidence=TextEvidence(conflicting_fields=("对方账号",)),
    )

    status, reason = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.PENDING_REVIEW
    assert "候选歧义" in reason
    assert "对方账号" in reason


def _std_df(rows):
    records = []
    for date, amount, summary, original_idx in rows:
        records.append(
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": summary,
                "aux_text_fields": {"摘要": summary},
                "original_idx": original_idx,
                "original_file_row": original_idx + 1,
                "voucher_no": "",
            }
        )
    return pd.DataFrame(records)


def test_连续日记账摘要整组可对应单笔银行流水():
    bank = _std_df([("2026-01-02", 300, "项目甲回款", 1)])
    journal = _std_df(
        [
            ("2026-01-01", 100, "项目甲回款", 8),
            ("2026-01-02", 200, "项目甲回款", 9),
        ]
    )
    matcher = Matcher(bank, journal, MatcherConfig(dfs_date_window=3))

    matcher.run()

    candidate = next(
        item
        for item in matcher.candidates
        if item.match_type == "continuous_summary_group"
    )
    assert candidate.match_type == "continuous_summary_group"
    assert candidate.match_stage == "连续摘要整组"
    assert candidate.bank_idxs == (0,)
    assert candidate.journal_idxs == (0, 1)
    assert candidate.evidence["source_side"] == "journal"
    assert candidate.evidence["source_rows"] == (8, 9)


def test_连续银行流水摘要整组也可对应单笔日记账():
    bank = _std_df(
        [
            ("2026-02-03", -40, "支付项目乙", 3),
            ("2026-02-04", -60, "支付项目乙", 4),
        ]
    )
    journal = _std_df([("2026-02-04", -100, "支付项目乙", 20)])
    matcher = Matcher(bank, journal, MatcherConfig(dfs_date_window=3))

    matcher.match_continuous_summary_groups()

    candidate = matcher.candidates[0]
    assert candidate.bank_idxs == (0, 1)
    assert candidate.journal_idxs == (0,)
    assert candidate.evidence["source_side"] == "bank"


def test_原表行号不连续时不得拼成连续摘要整组():
    bank = _std_df([("2026-03-01", 300, "项目丙", 1)])
    journal = _std_df(
        [
            ("2026-03-01", 100, "项目丙", 5),
            ("2026-03-01", 200, "项目丙", 7),
        ]
    )
    matcher = Matcher(bank, journal, MatcherConfig())

    matcher.match_continuous_summary_groups()

    assert matcher.candidates == []


def test_收入支出抵销后总额相同也不得整组匹配():
    bank = _std_df([("2026-04-01", 0, "内部调账", 1)])
    journal = _std_df(
        [
            ("2026-04-01", 100, "内部调账", 10),
            ("2026-04-01", -100, "内部调账", 11),
        ]
    )
    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(allow_zero_match=True),
    )

    matcher.match_continuous_summary_groups()

    assert matcher.candidates == []
