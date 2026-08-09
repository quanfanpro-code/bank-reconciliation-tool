"""匹配可靠性补强测试。"""

from data_structures import (
    GroupMetrics,
    MatchCandidate,
    MatcherConfig,
    ProcessingStatus,
    TextEvidence,
)
from matching_policy import route_candidate


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
