"""自动化优先核对策略测试。"""

from decimal import Decimal

import data_structures as ds
import pandas as pd
from balance import BalanceRecalculator, BalanceReconciler
from matcher import Matcher
from matching_policy import (
    apply_monthly_difference_pools,
    build_group_metrics,
    risk_level_for,
    route_candidate,
)
from precision_engine import PrecisionEngine


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


def _candidate(
    *,
    bank_amounts,
    journal_amounts,
    ambiguous=False,
):
    return ds.MatchCandidate(
        candidate_id="C1",
        bank_idxs=(0,),
        journal_idxs=(0,),
        match_type="测试候选",
        match_stage="测试",
        metrics=build_group_metrics(
            [PrecisionEngine.to_integer_li(value) for value in bank_amounts],
            [PrecisionEngine.to_integer_li(value) for value in journal_amounts],
        ),
        is_ambiguous=ambiguous,
    )


def test_重大歧义自动形成高风险疑点而不等待人工():
    candidate = _candidate(
        bank_amounts=[Decimal("120000")],
        journal_amounts=[Decimal("120000")],
        ambiguous=True,
    )

    status, risk, reason = route_candidate(candidate, ds.MatcherConfig())

    assert status is ds.ProcessingStatus.FLAGGED
    assert risk is ds.RiskLevel.HIGH
    assert "候选歧义" in reason


def test_小额差异自动归集为低风险():
    candidate = _candidate(
        bank_amounts=[Decimal("10000")],
        journal_amounts=[Decimal("9900")],
    )

    status, risk, reason = route_candidate(candidate, ds.MatcherConfig())

    assert status is ds.ProcessingStatus.AUTO_CLASSIFIED
    assert risk is ds.RiskLevel.LOW
    assert "差异" in reason


def test_月度累计超限只升级风险而不改成人工状态():
    candidates = []
    for index in range(21):
        candidate = _candidate(
            bank_amounts=[Decimal("5000")],
            journal_amounts=[Decimal("0")],
        )
        candidate.candidate_id = f"C{index:02d}"
        candidate.bank_dates = (ds.pd.Timestamp(f"2026-01-{index + 1:02d}"),)
        candidate.journal_dates = candidate.bank_dates
        candidate.processing_status = ds.ProcessingStatus.AUTO_CLASSIFIED
        candidate.risk_level = ds.RiskLevel.LOW
        candidates.append(candidate)

    pool = apply_monthly_difference_pools(
        candidates,
        ds.MatcherConfig(),
    )[0]

    assert pool.risk_level is ds.RiskLevel.HIGH
    assert all(
        candidate.processing_status is ds.ProcessingStatus.AUTO_CLASSIFIED
        for candidate in candidates
    )
    assert all(
        candidate.risk_level is ds.RiskLevel.HIGH
        for candidate in candidates
    )


def test_唯一一对多组合自动确认():
    bank = _std_df([("2026-01-02", 100, "项目回款", 1)])
    journal = _std_df(
        [
            ("2026-01-02", 40, "项目回款", 10),
            ("2026-01-03", 60, "项目回款", 11),
        ]
    )
    matcher = Matcher(bank, journal, ds.MatcherConfig(dfs_date_window=3))

    matcher._collecting_candidates = True
    matcher.match_dfs_combinations()
    matcher._collecting_candidates = False
    matcher._commit_selected_candidates()

    assert len(matcher.selected_candidates) == 1
    selected = matcher.selected_candidates[0]
    assert selected.bank_idxs == (0,)
    assert selected.journal_idxs == (0, 1)
    assert selected.processing_status is ds.ProcessingStatus.AUTO_CONFIRMED


def test_多解但整组金额闭合时自动输出整组勾稽一致():
    bank = _std_df(
        [
            ("2026-01-02", 100, "项目回款", 1),
            ("2026-01-02", 100, "项目回款", 2),
        ]
    )
    journal = _std_df(
        [
            ("2026-01-02", 40, "项目回款", 10),
            ("2026-01-02", 60, "项目回款", 11),
            ("2026-01-02", 30, "项目回款", 12),
            ("2026-01-02", 70, "项目回款", 13),
        ]
    )
    matcher = Matcher(bank, journal, ds.MatcherConfig(dfs_date_window=3))

    matcher._collecting_candidates = True
    matcher.match_dfs_combinations()
    matcher._collecting_candidates = False
    matcher._commit_selected_candidates()

    assert len(matcher.selected_candidates) == 1
    selected = matcher.selected_candidates[0]
    assert selected.bank_idxs == (0, 1)
    assert selected.journal_idxs == (0, 1, 2, 3)
    assert selected.processing_status is ds.ProcessingStatus.GROUP_RECONCILED
    assert selected.evidence["closed_group_fallback"] is True


def test_没有有效余额时不从零伪造余额轨迹():
    frame = _std_df([("2026-01-02", 100, "项目回款", 1)])

    assert BalanceRecalculator().recalculate(frame) == []


def test_余额列全部为空时也不从零伪造余额轨迹():
    frame = _std_df([("2026-01-02", 70, "项目回款", 1)])
    frame["balance"] = [None]

    assert BalanceRecalculator().recalculate(frame) == []


def test_零影响且无定性风险时风险等级为正常():
    assert risk_level_for(0, ds.MatcherConfig()) is ds.RiskLevel.NORMAL


def test_双方余额完全相等时不生成余额差异():
    daily = ds.DailyBalance(
        date=pd.Timestamp("2026-01-02"),
        income=Decimal("100"),
        expense=Decimal("0"),
        net=Decimal("100"),
        balance=Decimal("1000"),
        prev_balance=Decimal("900"),
    )

    assert BalanceReconciler([daily], [daily]).generate_diff_report() == []
