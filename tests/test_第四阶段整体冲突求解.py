from data_structures import GroupMetrics, MatchCandidate, ScoreBreakdown
from matcher import select_non_conflicting_candidates


def _candidate(candidate_id, bank, journal, score=80, *, full=False, strength=0):
    return MatchCandidate(
        candidate_id=candidate_id,
        bank_idxs=tuple(bank),
        journal_idxs=tuple(journal),
        match_type="test",
        match_stage="test",
        metrics=GroupMetrics(100, 100, 100, 100, 100, 0, 0, 0, 0, 0),
        scores=ScoreBreakdown(amount=40, date=20, text=max(0, score - 70), structure=10),
        evidence={
            "resolves_full_group": full,
            "represents_full_observed_group": full,
            "complete_business_id": full,
            "business_strength": strength,
        },
        stable_key=candidate_id,
        composition_key=candidate_id,
    )


def test_整体求解能避开贪心首项并选择覆盖更多原始行的组合():
    candidates = [
        _candidate("A", (0,), (0,), 99),
        _candidate("B", (0,), (1, 2), 90),
        _candidate("C", (1, 2), (0,), 90),
    ]
    diagnostics = {}

    selected = select_non_conflicting_candidates(candidates, diagnostics=diagnostics)

    assert {item.candidate_id for item in selected} == {"B", "C"}
    assert diagnostics["exact_components"] == 1
    assert diagnostics["fallback_components"] == 0


def test_完整强证据业务组优先于拆成多个弱关系():
    full = _candidate("FULL", tuple(range(35)), (0,), 80, full=True, strength=4)
    weak = [
        _candidate(f"W{index:02d}", (index,), (index + 1,), 99)
        for index in range(20)
    ]

    selected = select_non_conflicting_candidates([*weak, full])

    assert [item.candidate_id for item in selected] == ["FULL"]


def test_超大竞争组使用确定性降级且如实披露():
    candidates = [
        _candidate(f"C{index:02d}", (0,), (index,), 80 + index % 10)
        for index in range(30)
    ]
    first_diagnostics = {}
    second_diagnostics = {}

    first = select_non_conflicting_candidates(candidates, diagnostics=first_diagnostics, exact_component_limit=18)
    second = select_non_conflicting_candidates(list(reversed(candidates)), diagnostics=second_diagnostics, exact_component_limit=18)

    assert [item.candidate_id for item in first] == [item.candidate_id for item in second]
    assert first_diagnostics["fallback_components"] == 1
    assert first_diagnostics["search_fully_exhausted"] is False


def test_收到中止信号时不再穷举并披露中止():
    candidates = [
        _candidate(f"C{index:02d}", (index % 3,), (index % 4,), 80)
        for index in range(15)
    ]
    diagnostics = {}

    select_non_conflicting_candidates(
        candidates,
        diagnostics=diagnostics,
        should_stop=lambda: True,
    )

    assert diagnostics["stopped"] is True
    assert diagnostics["search_fully_exhausted"] is False
