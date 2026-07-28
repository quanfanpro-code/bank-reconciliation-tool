"""统一候选评分与分流使用的纯业务规则。"""

from collections.abc import Mapping, Sequence
from difflib import SequenceMatcher
import re
import unicodedata

from data_structures import (
    DifferenceComponent,
    DifferencePoolResult,
    DifferencePoolType,
    GroupMetrics,
    MatchCandidate,
    MatcherConfig,
    ProcessingStatus,
    ScoreBreakdown,
    TextEvidence,
)
from precision_engine import PrecisionEngine


def _normalize_text(value: object) -> str:
    """统一全半角、大小写和标点，但保留中文、字母与数字。"""
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).casefold().strip()
    if text in {"", "nan", "none", "null"}:
        return ""
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", text)


def _text_similarity(left: str, right: str) -> int:
    """综合字符顺序和相邻字符重合度，返回零到一百分。"""
    if not left or not right:
        return 0
    if left == right:
        return 100
    sequence_score = SequenceMatcher(None, left, right, autojunk=False).ratio()
    left_pairs = {left[index:index + 2] for index in range(len(left) - 1)}
    right_pairs = {right[index:index + 2] for index in range(len(right) - 1)}
    pair_score = 0.0
    if left_pairs and right_pairs:
        pair_score = (
            2 * len(left_pairs & right_pairs) / (len(left_pairs) + len(right_pairs))
        )
    return max(0, min(100, round(max(sequence_score, pair_score) * 100)))


def _critical_field_category(label: str) -> str:
    """把同义的关键账户字段归入同一个冲突检查类别。"""
    compact = _normalize_text(label)
    if "户名" in compact or "客户名称" in compact or "对方名称" in compact:
        return "对方户名"
    if "账号" in compact or "帐号" in compact:
        return "账号"
    if "卡号" in compact:
        return "卡号"
    return ""


def _find_text_conflicts(
    bank_fields: Mapping[str, object],
    journal_fields: Mapping[str, object],
) -> tuple[str, ...]:
    """仅对两侧都存在的关键账户字段识别明确矛盾。"""
    bank_values: dict[str, set[str]] = {}
    journal_values: dict[str, set[str]] = {}
    for label, value in bank_fields.items():
        category = _critical_field_category(label)
        normalized = _normalize_text(value)
        if category and normalized:
            bank_values.setdefault(category, set()).add(normalized)
    for label, value in journal_fields.items():
        category = _critical_field_category(label)
        normalized = _normalize_text(value)
        if category and normalized:
            journal_values.setdefault(category, set()).add(normalized)
    conflicts = [
        category
        for category in bank_values.keys() & journal_values.keys()
        if bank_values[category].isdisjoint(journal_values[category])
    ]
    return tuple(sorted(conflicts))


def build_sensitive_field_signals(
    bank_fields: Mapping[str, object],
    journal_fields: Mapping[str, object],
) -> tuple[str, ...]:
    """只输出敏感账户字段的一致性结论，绝不携带字段原值。"""
    def collect(
        fields: Mapping[str, object],
    ) -> dict[str, set[str]]:
        values: dict[str, set[str]] = {}
        for label, value in fields.items():
            category = _critical_field_category(str(label))
            normalized = _normalize_text(value)
            if category and normalized:
                values.setdefault(category, set()).add(normalized)
        return values

    bank_values = collect(bank_fields)
    journal_values = collect(journal_fields)
    signals = []
    for category in sorted(bank_values.keys() & journal_values.keys()):
        state = (
            "不同"
            if bank_values[category].isdisjoint(journal_values[category])
            else "相同"
        )
        signals.append(f"{category}:{state}")
    return tuple(signals)


def score_text_fields(
    bank_fields: Mapping[str, object],
    journal_fields: Mapping[str, object],
    *,
    llm_semantic_score: int | None = None,
) -> TextEvidence:
    """逐字段比较文本，大模型只能补充这一项证据。"""
    field_scores: dict[str, int] = {}
    for bank_label, bank_value in bank_fields.items():
        bank_text = _normalize_text(bank_value)
        if not bank_text:
            continue
        for journal_label, journal_value in journal_fields.items():
            journal_text = _normalize_text(journal_value)
            if not journal_text:
                continue
            label = f"银行.{bank_label} ↔ 日记账.{journal_label}"
            field_scores[label] = _text_similarity(bank_text, journal_text)

    ranked = sorted(field_scores.items(), key=lambda item: (-item[1], item[0]))
    if ranked:
        top_scores = [score for _, score in ranked[:3]]
        local_score = round(0.6 * top_scores[0] + 0.4 * sum(top_scores) / len(top_scores))
    else:
        local_score = 0

    model_score = None
    if llm_semantic_score is not None:
        model_score = max(0, min(100, round(llm_semantic_score)))
        combined_score = round(local_score * 0.4 + model_score * 0.6)
    else:
        combined_score = local_score

    supporting = tuple(label for label, score in ranked if score >= 60)
    conflicts = _find_text_conflicts(bank_fields, journal_fields)
    return TextEvidence(
        field_scores=field_scores,
        supporting_fields=supporting,
        conflicting_fields=conflicts,
        local_score=local_score,
        model_score=model_score,
        combined_score=combined_score,
        score=max(0, min(30, round(combined_score * 0.3))),
    )


def merge_labeled_fields(
    field_groups: Sequence[Mapping[str, object]],
) -> dict[str, str]:
    """按字段名合并多笔记录，绝不把不同字段混成无标签长文本。"""
    merged: dict[str, list[str]] = {}
    for fields in field_groups:
        if not isinstance(fields, dict):
            continue
        for label, value in fields.items():
            normalized = _normalize_text(value)
            if not normalized:
                continue
            display_value = str(value).strip()
            values = merged.setdefault(str(label), [])
            if display_value not in values:
                values.append(display_value)
    return {
        label: "；".join(values)
        for label, values in sorted(merged.items())
    }


def _amount_score(candidate: MatchCandidate, config: MatcherConfig) -> int:
    """完全一致四十分；差异越大，金额证据越弱。"""
    difference = candidate.metrics.total_diff_li
    if difference == 0:
        return 40
    trivial_li = max(
        1,
        PrecisionEngine.to_integer_li(config.clearly_trivial_threshold),
    )
    if difference <= trivial_li:
        return max(20, round(40 - 20 * difference / trivial_li))
    group_amount = max(1, candidate.metrics.group_amount_li)
    relative_difference = min(1.0, difference / group_amount)
    return max(0, round(20 * (1 - relative_difference)))


def _date_score(candidate: MatchCandidate, config: MatcherConfig) -> int:
    """同日十五分，到达日期容差边界时降为零分。"""
    span = max(0, candidate.date_span_days)
    tolerance = max(0, config.tolerance_days)
    if span == 0:
        return 15
    if tolerance == 0 or span >= tolerance:
        return 0
    return max(0, min(15, round(15 * (1 - span / tolerance))))


def _structure_score(
    candidate: MatchCandidate,
    text_evidence: TextEvidence,
) -> int:
    """匹配关系越复杂、歧义或冲突越多，结构分越低。"""
    bank_count = len(candidate.bank_idxs)
    journal_count = len(candidate.journal_idxs)
    if bank_count == 1 and journal_count == 1:
        score = 15
    elif min(bank_count, journal_count) == 1:
        score = 12
    else:
        score = 9
    if candidate.is_ambiguous:
        score -= 5
    if candidate.is_cross_month_many_to_many:
        score -= 5
    score -= min(6, len(text_evidence.conflicting_fields) * 3)
    if candidate.rule_matched and candidate.evidence.get(
        "structure_bonus",
        True,
    ):
        score += 6
    return max(0, min(15, score))


def _cap_ambiguous_without_text(scores: ScoreBreakdown) -> ScoreBreakdown:
    """无文本支持的歧义候选不能仅靠金额和日期越过自动确认线。"""
    if scores.total <= 69:
        return scores
    excess = scores.total - 69
    return ScoreBreakdown(
        amount=scores.amount,
        date=scores.date,
        text=scores.text,
        structure=max(0, scores.structure - excess),
    )


def score_candidate(
    candidate: MatchCandidate,
    config: MatcherConfig,
    bank_fields: Mapping[str, object],
    journal_fields: Mapping[str, object],
    *,
    llm_semantic_score: int | None = None,
) -> ScoreBreakdown:
    """所有匹配路径统一调用这一处，得到唯一综合可信度。"""
    text_evidence = score_text_fields(
        bank_fields,
        journal_fields,
        llm_semantic_score=llm_semantic_score,
    )
    scores = ScoreBreakdown(
        amount=_amount_score(candidate, config),
        date=_date_score(candidate, config),
        text=text_evidence.score,
        structure=_structure_score(candidate, text_evidence),
    )
    if candidate.is_ambiguous and text_evidence.score == 0:
        scores = _cap_ambiguous_without_text(scores)
    candidate.text_evidence = text_evidence
    candidate.scores = scores
    return scores


def build_group_metrics(
    bank_amounts_li: Sequence[int],
    journal_amounts_li: Sequence[int],
) -> GroupMetrics:
    """分别计算两侧收入、支出和绝对金额，不允许正负抵销。"""
    bank_income = sum(value for value in bank_amounts_li if value > 0)
    journal_income = sum(value for value in journal_amounts_li if value > 0)
    bank_expense = sum(abs(value) for value in bank_amounts_li if value < 0)
    journal_expense = sum(abs(value) for value in journal_amounts_li if value < 0)
    bank_gross = sum(abs(value) for value in bank_amounts_li)
    journal_gross = sum(abs(value) for value in journal_amounts_li)
    income_diff = abs(bank_income - journal_income)
    expense_diff = abs(bank_expense - journal_expense)
    return GroupMetrics(
        bank_gross_li=bank_gross,
        journal_gross_li=journal_gross,
        group_amount_li=max(bank_gross, journal_gross),
        bank_income_li=bank_income,
        journal_income_li=journal_income,
        bank_expense_li=bank_expense,
        journal_expense_li=journal_expense,
        income_diff_li=income_diff,
        expense_diff_li=expense_diff,
        total_diff_li=income_diff + expense_diff,
    )


def route_candidate(
    candidate: MatchCandidate,
    config: MatcherConfig,
) -> tuple[ProcessingStatus, str]:
    """按更严格规则优先确定候选的最终处理状态。"""
    performance_li = PrecisionEngine.to_integer_li(config.performance_materiality)
    trivial_li = PrecisionEngine.to_integer_li(config.clearly_trivial_threshold)

    if candidate.metrics.group_amount_li > performance_li:
        return (
            ProcessingStatus.PENDING_REVIEW,
            "匹配组金额超过实际执行重要性水平",
        )
    if candidate.is_cross_month_many_to_many:
        return ProcessingStatus.PENDING_REVIEW, "跨月多对多必须人工复核"
    if candidate.metrics.total_diff_li == 0:
        return ProcessingStatus.AUTO_CONFIRMED, "金额完全一致"
    if candidate.metrics.total_diff_li <= trivial_li:
        return ProcessingStatus.AUTO_CONFIRMED, "明显微小错报自动处理"
    if candidate.scores.total >= config.auto_confirm_score:
        return ProcessingStatus.AUTO_CONFIRMED, "综合可信度达到自动确认门槛"
    return ProcessingStatus.PENDING_REVIEW, "综合可信度未达到自动确认门槛"


def bucket_distribution(amounts_li: Sequence[int]) -> tuple[int, ...]:
    """按 100、1,000、10,000、100,000 元边界统计绝对金额分布。"""
    boundaries = [
        0,
        PrecisionEngine.to_integer_li("100"),
        PrecisionEngine.to_integer_li("1000"),
        PrecisionEngine.to_integer_li("10000"),
        PrecisionEngine.to_integer_li("100000"),
        float("inf"),
    ]
    distribution = [0] * (len(boundaries) - 1)
    for amount in amounts_li:
        absolute = abs(amount)
        for index in range(len(boundaries) - 1):
            if boundaries[index] <= absolute < boundaries[index + 1]:
                distribution[index] += 1
                break
    return tuple(distribution)


def _component_month(
    candidate: MatchCandidate,
    *,
    source: str,
) -> str:
    dates = (
        candidate.bank_dates if source == "bank" else candidate.journal_dates
    )
    if not dates:
        dates = candidate.bank_dates + candidate.journal_dates
    if not dates:
        return "未知月份"
    return min(dates).strftime("%Y-%m")


def build_difference_components(
    candidate: MatchCandidate,
) -> list[DifferenceComponent]:
    """按来源和收支方向拆分差异，正负差异绝不抵销。"""
    metrics = candidate.metrics
    components: list[DifferenceComponent] = []

    def add(
        pool_type: DifferencePoolType,
        difference: int,
        source: str,
    ) -> None:
        if difference <= 0:
            return
        components.append(
            DifferenceComponent(
                month=_component_month(candidate, source=source),
                pool_type=pool_type,
                candidate_id=candidate.candidate_id,
                diff_li=difference,
            )
        )

    if metrics.bank_income_li > metrics.journal_income_li:
        add(
            DifferencePoolType.BANK_ONLY_INCOME,
            metrics.bank_income_li - metrics.journal_income_li,
            "bank",
        )
    elif metrics.journal_income_li > metrics.bank_income_li:
        add(
            DifferencePoolType.JOURNAL_ONLY_INCOME,
            metrics.journal_income_li - metrics.bank_income_li,
            "journal",
        )

    if metrics.bank_expense_li > metrics.journal_expense_li:
        add(
            DifferencePoolType.BANK_ONLY_EXPENSE,
            metrics.bank_expense_li - metrics.journal_expense_li,
            "bank",
        )
    elif metrics.journal_expense_li > metrics.bank_expense_li:
        add(
            DifferencePoolType.JOURNAL_ONLY_EXPENSE,
            metrics.journal_expense_li - metrics.bank_expense_li,
            "journal",
        )
    return components


def apply_monthly_difference_pools(
    candidates: Sequence[MatchCandidate],
    config: MatcherConfig,
) -> list[DifferencePoolResult]:
    """按自然月和四种来源方向累计明显微小错报。"""
    trivial_li = PrecisionEngine.to_integer_li(
        config.clearly_trivial_threshold
    )
    performance_li = PrecisionEngine.to_integer_li(
        config.performance_materiality
    )
    pools: dict[
        tuple[str, DifferencePoolType],
        DifferencePoolResult,
    ] = {}
    candidate_by_id = {
        candidate.candidate_id: candidate for candidate in candidates
    }
    for candidate in candidates:
        candidate.evidence.setdefault("included_in_pool_review", False)
        if candidate.metrics.total_diff_li <= 0:
            continue
        if candidate.metrics.total_diff_li > trivial_li:
            continue
        if candidate.metrics.group_amount_li > performance_li:
            continue
        if candidate.is_cross_month_many_to_many:
            continue
        for component in build_difference_components(candidate):
            key = (component.month, component.pool_type)
            if key not in pools:
                pools[key] = DifferencePoolResult(
                    pool_id=(
                        f"POOL|{component.month}|"
                        f"{component.pool_type.value}"
                    ),
                    month=component.month,
                    pool_type=component.pool_type,
                )
            pool = pools[key]
            pool.total_diff_li += abs(component.diff_li)
            pool.components.append(component)

    results = sorted(
        pools.values(),
        key=lambda pool: (pool.month, pool.pool_type.value),
    )
    for pool in results:
        pool.exceeds_performance_materiality = (
            pool.total_diff_li > performance_li
        )
        if pool.exceeds_performance_materiality:
            pool.processing_status = ProcessingStatus.PENDING_REVIEW
            pool.processing_reason = "月度累计超出实际执行重要性水平"
            for component in pool.components:
                component.included_in_pool_review = True
                matched_candidate = candidate_by_id.get(
                    component.candidate_id
                )
                if matched_candidate is not None:
                    matched_candidate.evidence[
                        "included_in_pool_review"
                    ] = True
                    matched_candidate.processing_status = (
                        ProcessingStatus.PENDING_REVIEW
                    )
                    matched_candidate.processing_reason = (
                        "纳入月度差异池整池复核"
                    )
                    pool_ids = matched_candidate.evidence.setdefault(
                        "difference_pool_ids",
                        [],
                    )
                    if pool.pool_id not in pool_ids:
                        pool_ids.append(pool.pool_id)
        else:
            pool.processing_status = ProcessingStatus.AUTO_CONFIRMED
            pool.processing_reason = "月度累计未超过实际执行重要性水平"
    return results
