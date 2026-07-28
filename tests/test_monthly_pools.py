from decimal import Decimal

import pandas as pd

from data_structures import (
    DifferencePoolType,
    MatchCandidate,
    MatcherConfig,
    ProcessingStatus,
)
from matcher import Matcher
from matching_policy import (
    apply_monthly_difference_pools,
    build_difference_components,
    build_group_metrics,
)
from precision_engine import PrecisionEngine


def _candidate(
    candidate_id,
    date,
    bank_amounts,
    journal_amounts,
):
    timestamp = pd.Timestamp(date)
    return MatchCandidate(
        candidate_id=candidate_id,
        bank_idxs=(0,),
        journal_idxs=(0,),
        match_type="金额差异",
        match_stage="测试",
        metrics=build_group_metrics(
            [PrecisionEngine.to_integer_li(value) for value in bank_amounts],
            [
                PrecisionEngine.to_integer_li(value)
                for value in journal_amounts
            ],
        ),
        bank_dates=(timestamp,),
        journal_dates=(timestamp,),
    )


def test_同月四类差异分别累计且收入支出不抵销():
    candidates = [
        _candidate("C1", "2026-01-05", [3000], [0]),
        _candidate("C2", "2026-01-06", [-3000], [0]),
        _candidate("C3", "2026-01-07", [0], [2000]),
        _candidate("C4", "2026-01-08", [0], [-1000]),
    ]

    pools = apply_monthly_difference_pools(
        candidates,
        MatcherConfig(),
    )
    totals = {
        (pool.month, pool.pool_type): pool.total_diff_li
        for pool in pools
    }

    assert totals[
        ("2026-01", DifferencePoolType.BANK_ONLY_INCOME)
    ] == PrecisionEngine.to_integer_li("3000")
    assert totals[
        ("2026-01", DifferencePoolType.BANK_ONLY_EXPENSE)
    ] == PrecisionEngine.to_integer_li("3000")
    assert totals[
        ("2026-01", DifferencePoolType.JOURNAL_ONLY_INCOME)
    ] == PrecisionEngine.to_integer_li("2000")
    assert totals[
        ("2026-01", DifferencePoolType.JOURNAL_ONLY_EXPENSE)
    ] == PrecisionEngine.to_integer_li("1000")


def test_单池累计超过实际执行重要性时只形成一项整池待复核():
    candidates = [
        _candidate(
            f"C{index:02d}",
            f"2026-01-{index + 1:02d}",
            [5000],
            [0],
        )
        for index in range(21)
    ]
    for candidate in candidates:
        candidate.processing_status = ProcessingStatus.AUTO_CONFIRMED
        candidate.processing_reason = "明显微小错报自动处理"

    pools = apply_monthly_difference_pools(
        candidates,
        MatcherConfig(),
    )
    over_limit = [
        pool
        for pool in pools
        if pool.exceeds_performance_materiality
    ]

    assert len(over_limit) == 1
    assert (
        over_limit[0].processing_status
        is ProcessingStatus.PENDING_REVIEW
    )
    assert len(over_limit[0].components) == 21
    assert all(
        component.included_in_pool_review
        for component in over_limit[0].components
    )
    assert all(
        candidate.evidence["included_in_pool_review"] is True
        for candidate in candidates
    )
    assert all(
        candidate.processing_status is ProcessingStatus.PENDING_REVIEW
        for candidate in candidates
    )
    assert all(
        candidate.processing_reason == "纳入月度差异池整池复核"
        for candidate in candidates
    )


def test_单池累计刚好等于实际执行重要性水平不算超过():
    candidates = [
        _candidate(
            f"C{index:02d}",
            f"2026-01-{index + 1:02d}",
            [5000],
            [0],
        )
        for index in range(20)
    ]

    pool = apply_monthly_difference_pools(
        candidates,
        MatcherConfig(),
    )[0]

    assert pool.total_diff_li == PrecisionEngine.to_integer_li("100000")
    assert pool.exceeds_performance_materiality is False
    assert pool.processing_status is ProcessingStatus.AUTO_CONFIRMED


def test_不同自然月的同类差异不能合并累计():
    candidates = [
        _candidate("C1", "2026-01-31", [4000], [0]),
        _candidate("C2", "2026-02-01", [3000], [0]),
    ]

    pools = apply_monthly_difference_pools(
        candidates,
        MatcherConfig(),
    )

    assert [(pool.month, pool.total_diff_li) for pool in pools] == [
        ("2026-01", PrecisionEngine.to_integer_li("4000")),
        ("2026-02", PrecisionEngine.to_integer_li("3000")),
    ]


def test_单个候选同时存在收入和支出差异时拆成两个组成():
    candidate = _candidate(
        "C1",
        "2026-01-05",
        [2000, -2000],
        [0],
    )

    components = build_difference_components(candidate)

    assert len(components) == 2
    assert {item.pool_type for item in components} == {
        DifferencePoolType.BANK_ONLY_INCOME,
        DifferencePoolType.BANK_ONLY_EXPENSE,
    }


def _std_df(date, amount, summary):
    return pd.DataFrame(
        [
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": summary,
                "aux_text_fields": {"摘要": summary},
                "original_idx": 1,
                "original_file_row": 2,
                "voucher_no": "",
            }
        ]
    )


def test_真实匹配会生成金额差异候选并进入对应月度池():
    matcher = Matcher(
        _std_df("2026-01-05", 1000, "甲公司设备款"),
        _std_df("2026-01-05", 900, "甲公司设备款"),
        MatcherConfig(),
    )

    matcher.run()

    assert matcher.selected_candidates[0].match_type == "amount_difference"
    assert (
        matcher.selected_candidates[0].processing_status
        is ProcessingStatus.AUTO_CONFIRMED
    )
    assert len(matcher.difference_pools) == 1
    assert (
        matcher.difference_pools[0].pool_type
        is DifferencePoolType.BANK_ONLY_INCOME
    )
    assert matcher.difference_pools[0].total_diff_li == (
        PrecisionEngine.to_integer_li("100")
    )
