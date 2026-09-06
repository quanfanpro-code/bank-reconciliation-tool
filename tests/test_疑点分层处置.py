"""疑点分层处置测试：明显微小零差额自动确认 + 中风险等距抽样。"""

from decimal import Decimal

import pandas as pd

import data_structures as ds
from data_structures import (
    MatcherConfig,
    ProcessingStatus,
    RiskLevel,
    ScoreBreakdown,
    TextEvidence,
)
from matcher import Matcher
from matching_policy import (
    apply_medium_risk_sampling,
    apply_monthly_difference_pools,
    build_group_metrics,
    route_candidate,
)
from precision_engine import PrecisionEngine
from reporter import Reporter


def _candidate(
    *,
    bank_amounts,
    journal_amounts,
    ambiguous=False,
    score=0,
    business_strength=0,
    date="2026-01-02",
    candidate_id="C1",
):
    candidate = ds.MatchCandidate(
        candidate_id=candidate_id,
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
    candidate.scores = ScoreBreakdown(
        amount=score, date=score, text=score, structure=score
    )
    candidate.bank_dates = (pd.Timestamp(date),)
    candidate.journal_dates = (pd.Timestamp(date),)
    candidate.evidence["business_strength"] = business_strength
    return candidate


def _medium_flagged(index, date, status=ProcessingStatus.FLAGGED):
    candidate = _candidate(
        bank_amounts=[Decimal("6000")],
        journal_amounts=[Decimal("6000")],
        ambiguous=True,
        date=date,
        candidate_id=f"M{index:02d}",
    )
    candidate.processing_status = status
    candidate.risk_level = RiskLevel.MEDIUM
    return candidate


def _high_flagged(index):
    candidate = _candidate(
        bank_amounts=[Decimal("200000")],
        journal_amounts=[Decimal("200000")],
        ambiguous=True,
        candidate_id=f"H{index:02d}",
    )
    candidate.processing_status = ProcessingStatus.FLAGGED
    candidate.risk_level = RiskLevel.HIGH
    return candidate


def _report_records(rows):
    records = []
    for original_idx, (date, amount, summary, source) in enumerate(rows, start=1):
        records.append(
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": summary,
                "aux_text_fields": {"摘要": summary},
                "balance": None,
                "voucher_no": f"记-{original_idx:03d}" if source == "journal" else "",
                "source": source,
                "original_idx": original_idx,
                "original_file_row": original_idx + 1,
            }
        )
    return pd.DataFrame(records)


def test_明显微小零差额弱证据匹配自动确认():
    candidate = _candidate(
        bank_amounts=[Decimal("3000")],
        journal_amounts=[Decimal("3000")],
        ambiguous=True,
    )

    status, risk, reason = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.AUTO_CONFIRMED
    assert risk is RiskLevel.NORMAL
    assert "金额明显微小自动确认" in reason
    assert "候选歧义" in reason
    assert "业务依据不足" in reason


def test_超过明显微小临界值的弱证据匹配维持疑点():
    candidate = _candidate(
        bank_amounts=[Decimal("6000")],
        journal_amounts=[Decimal("6000")],
        ambiguous=True,
    )

    status, risk, reason = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.FLAGGED
    assert risk is RiskLevel.MEDIUM
    assert "候选歧义" in reason


def test_明显微小但文字冲突维持疑点():
    candidate = _candidate(
        bank_amounts=[Decimal("3000")],
        journal_amounts=[Decimal("3000")],
    )
    candidate.text_evidence = TextEvidence(conflicting_fields=("对方户名",))

    status, risk, reason = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.FLAGGED
    assert "对方户名" in reason


def test_明显微小但在断档窗口内维持疑点():
    candidate = _candidate(
        bank_amounts=[Decimal("3000")],
        journal_amounts=[Decimal("3000")],
    )
    candidate.evidence["overall_scope_limited"] = True
    candidate.evidence["overall_control_reasons"] = ("余额在2026-07-01断档",)

    status, risk, reason = route_candidate(candidate, MatcherConfig())

    assert status is ProcessingStatus.FLAGGED
    assert risk is RiskLevel.UNKNOWN
    assert "总体资料尚未闭合" in reason


def test_中风险等距抽样样本量与高风险一致():
    mediums = [
        _medium_flagged(index, f"2026-01-{index + 1:02d}") for index in range(10)
    ]
    highs = [_high_flagged(index) for index in range(3)]

    stats = apply_medium_risk_sampling(list(reversed(mediums)) + highs)

    assert stats == {"medium_total": 10, "sampled": 3}
    sampled_ids = {
        candidate.candidate_id
        for candidate in mediums
        if candidate.evidence.get("medium_sampling") == "抽中待核查"
    }
    assert sampled_ids == {"M00", "M03", "M06"}
    for candidate in mediums:
        if candidate.candidate_id not in sampled_ids:
            assert candidate.evidence["medium_sampling"] == "未抽中留存备查"
    assert all("medium_sampling" not in candidate.evidence for candidate in highs)


def test_高风险为零时中风险全部留存备查():
    mediums = [
        _medium_flagged(index, f"2026-02-{index + 1:02d}") for index in range(3)
    ]
    mediums.append(
        _medium_flagged(3, "2026-02-04", status=ProcessingStatus.AUTO_CLASSIFIED)
    )

    stats = apply_medium_risk_sampling(mediums)

    assert stats == {"medium_total": 4, "sampled": 0}
    assert all(
        candidate.evidence["medium_sampling"] == "未抽中留存备查"
        for candidate in mediums
    )


def test_差异池升级的高风险计入抽样基数():
    pooled = []
    for index in range(21):
        candidate = _candidate(
            bank_amounts=[Decimal("5000")],
            journal_amounts=[Decimal("0")],
            candidate_id=f"P{index:02d}",
            date=f"2026-01-{index + 1:02d}",
        )
        candidate.processing_status = ProcessingStatus.AUTO_CLASSIFIED
        candidate.risk_level = RiskLevel.LOW
        pooled.append(candidate)
    apply_monthly_difference_pools(pooled, MatcherConfig())
    assert all(candidate.risk_level is RiskLevel.HIGH for candidate in pooled)

    mediums = [
        _medium_flagged(index, f"2026-03-{index + 1:02d}") for index in range(5)
    ]
    stats = apply_medium_risk_sampling(pooled + mediums)

    assert stats == {"medium_total": 5, "sampled": 5}
    assert all(
        candidate.evidence["medium_sampling"] == "抽中待核查"
        for candidate in mediums
    )


def test_中风险抽样结果写入疑点事项建议动作与总览():
    rows = []
    for index, amount in enumerate((6000, 7000, 8000, 9000)):
        date = f"2026-03-{index + 1:02d}"
        rows.append((date, amount, "服务费", "bank"))
        rows.append((date, amount, "服务费", "journal"))
        rows.append((date, amount, "服务费", "journal"))
    rows.append(("2026-04-01", 200000, "设备款", "bank"))
    rows.append(("2026-04-01", 200000, "设备款", "journal"))
    rows.append(("2026-04-01", 200000, "设备款", "journal"))
    bank = _report_records([row for row in rows if row[3] == "bank"])
    journal = _report_records([row for row in rows if row[3] == "journal"])
    matcher = Matcher(bank, journal, MatcherConfig())

    matcher.run()

    assert matcher.medium_sampling_stats == {"medium_total": 4, "sampled": 1}
    tables = Reporter(matcher).build_report_tables(MatcherConfig())
    issues = tables["疑点事项"]
    medium_rows = issues[issues["风险等级"] == "中风险"]
    assert len(medium_rows) == 4
    unsampled = medium_rows[
        medium_rows["建议动作"].str.contains("未抽中", na=False)
    ]
    sampled = medium_rows[
        ~medium_rows["建议动作"].str.contains("未抽中", na=False)
    ]
    assert len(unsampled) == 3
    assert len(sampled) == 1
    assert unsampled["判断依据"].str.contains("中风险等距抽样：未抽中").all()
    assert sampled["判断依据"].str.contains("中风险等距抽样：抽中").all()
    summary = dict(
        zip(tables["核对结论"]["项目"], tables["核对结论"]["数值"])
    )
    assert summary["中风险抽样总数"] == 4
    assert summary["中风险抽中待核查数"] == 1
