"""匹配可靠性补强测试。"""

from data_structures import (
    GroupMetrics,
    MatchCandidate,
    MatcherConfig,
    ProcessingStatus,
)
from matching_policy import route_candidate


def test_无风险且金额完全一致仍可自动确认():
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

    status, _ = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.AUTO_CONFIRMED
