"""把已有关系和两侧剩余记录归为完整复核事项，不改动匹配器结果。"""

from collections import defaultdict

import pandas as pd

from data_structures import ProcessingStatus, RiskLevel
from matching_policy import build_difference_components, build_group_metrics, risk_level_for
from precision_engine import PrecisionEngine


def build_review_items(matcher) -> list[dict]:
    """先补全原记录归属和月度累计，再对完整中风险总体等距抽样。"""
    items = []
    occupied = {"bank": set(), "journal": set()}
    frames = {"bank": matcher.bank, "journal": matcher.journal}
    candidates = list(getattr(matcher, "selected_candidates", []))
    by_candidate = {}
    pool_parts = defaultdict(list)
    risk_order = {risk.value: order for order, risk in enumerate(RiskLevel)}
    to_amount = PrecisionEngine.from_integer_li

    def make_item(identity, bank_idxs, journal_idxs, conclusion, risk, impact, difference, reason, **extra):
        indexes = {"bank": tuple(bank_idxs), "journal": tuple(journal_idxs)}
        amounts = {side: [int(frames[side].loc[index, "amount_decimal"]) for index in indexes[side]] for side in frames}
        days = [pd.Timestamp(frames[side].loc[index, "date"]) for side in frames for index in indexes[side]]
        metrics = build_group_metrics(amounts["bank"], amounts["journal"])
        item = {
            "事项编号": identity, "匹配ID": "", "银行索引": indexes["bank"], "日记账索引": indexes["journal"],
            "程序结论": conclusion, "风险等级": risk.value, "核对方式": "",
            "影响金额": to_amount(impact), "差异金额": to_amount(difference),
            "银行收入": to_amount(metrics.bank_income_li), "银行支出": to_amount(metrics.bank_expense_li),
            "日记账收入": to_amount(metrics.journal_income_li), "日记账支出": to_amount(metrics.journal_expense_li),
            "首次日期": min(days), "入选原因": "", "判断依据": reason,
            "跨期": len({day.strftime("%Y-%m") for day in days}) > 1,
            **extra,
        }
        for side in frames:
            if occupied[side].intersection(indexes[side]):
                raise ValueError("同一原始记录不能重复归入复核事项")
            occupied[side].update(indexes[side])
        items.append(item)
        return item

    for candidate in candidates:
        metrics = candidate.metrics
        relationship_risk = (
            candidate.is_ambiguous or candidate.is_cross_month_many_to_many
            or bool(candidate.text_evidence and candidate.text_evidence.conflicting_fields)
            or any(candidate.evidence.get(key) for key in (
                "total_only_without_boundary", "batch_boundary_uncertain", "overall_scope_limited"))
            or (candidate.evidence.get("represents_full_observed_group") and not candidate.evidence.get("resolves_full_group"))
        )
        if candidate.processing_status in {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}:
            impact = 0
        elif metrics.total_diff_li == 0 or relationship_risk:
            impact = metrics.group_amount_li
        else:
            impact = metrics.total_diff_li
        item = make_item(
            candidate.final_match_id, candidate.bank_idxs, candidate.journal_idxs,
            candidate.processing_status.value, candidate.risk_level, impact,
            metrics.total_diff_li, candidate.processing_reason,
            匹配ID=candidate.final_match_id, 候选ID=candidate.candidate_id,
        )
        by_candidate[candidate.candidate_id] = item

    # 使用原有差异组成和归月规则；只新增未对应原行，不改写原有匹配结果。
    existing_pools = getattr(matcher, "difference_pools", [])
    components = (
        [component for pool in existing_pools for component in pool.components]
        if existing_pools else
        [component for candidate in candidates if candidate.metrics.total_diff_li > 0
         for component in build_difference_components(candidate)]
    )
    for component in components:
        item = by_candidate.get(component.candidate_id)
        if item is not None:
            pool_parts[(component.month, component.pool_type.value)].append((item, component.diff_li))

    def remaining_item(side, indexes, event=None):
        frame = frames[side]
        indexes = tuple(indexes)
        rows = frame.loc[list(indexes)]
        amounts = [int(value) for value in rows["amount_decimal"]]
        impact = sum(abs(value) for value in amounts)
        control = getattr(matcher, "overall_control", None)
        days = [pd.Timestamp(value).normalize() for value in rows["date"]]
        hit_windows = [window for window in getattr(control, "affected_windows", ())
                       if any(window[0] <= day <= window[1] for day in days)]
        limited = bool(getattr(matcher, "overall_scope_limited", False)
                       or getattr(control, "scope_limited", False) or hit_windows)
        risk = risk_level_for(impact, matcher.config, unquantifiable=limited)
        original_rows = sorted(int(frame.loc[index].get("original_file_row", frame.loc[index].get("original_idx", index))) for index in indexes)
        side_name = "银行侧" if side == "bank" else "序时账侧"
        identity = f"{side_name}-原行" + "-".join(str(value) for value in original_rows)
        reason = "程序未找到足以建立跨双方对应关系的候选"
        if event:
            reason = f"{event.event_type}：{event.relationship_formula}；{event.evidence_basis}；尚未与另一侧建立对应"
        if limited:
            reasons = list(getattr(control, "reasons", ()))
            reasons.extend(f"余额断档窗口{start:%Y-%m-%d}至{end:%Y-%m-%d}" for start, end in hit_windows)
            reason += "；核对范围受限" + ("：" + "；".join(reasons) if reasons else "")
        item = make_item(
            identity, indexes if side == "bank" else (), indexes if side == "journal" else (),
            "同侧业务链" if event else f"{side_name}未对应", risk, impact, impact, reason,
        )
        if event:
            item["同侧业务链ID"] = event.event_id
            item["净额"] = to_amount(sum(amounts))
            item["跨期"] = item["跨期"] or event.is_cross_period
        for index in indexes:
            row = frame.loc[index]
            amount = int(row["amount_decimal"])
            if amount:
                direction = "收入" if amount > 0 else "支出"
                kind = ("银行已记公司未记-" if side == "bank" else "公司已记银行未记-") + direction
                pool_parts[(pd.Timestamp(row["date"]).strftime("%Y-%m"), kind)].append((item, abs(amount)))

    for event in getattr(matcher, "business_events", []):
        if event.source not in frames or not event.row_idxs:
            continue
        if all(index in frames[event.source].index and index not in occupied[event.source] for index in event.row_idxs):
            remaining_item(event.source, event.row_idxs, event)
    for side, frame in frames.items():
        for index in frame.index:
            if index not in occupied[side]:
                remaining_item(side, (int(index),))

    for (month, kind), parts in sorted(pool_parts.items()):
        total = sum(amount for _, amount in parts)
        pool_risk = risk_level_for(total, matcher.config)
        for item in {part[0]["事项编号"]: part[0] for part in parts}.values():
            if risk_order[pool_risk.value] > risk_order[item["风险等级"]]:
                item["风险等级"] = pool_risk.value
            if pool_risk is RiskLevel.HIGH:
                item["入选原因"] += f"{month}{kind}月度累计{to_amount(total)}元超过实际执行重要性水平；"

    for item in items:
        risk = item["风险等级"]
        if item["程序结论"] in {ProcessingStatus.AUTO_CONFIRMED.value, ProcessingStatus.GROUP_RECONCILED.value} and risk == RiskLevel.NORMAL.value:
            item["核对方式"] = "自动确认"
            item["入选原因"] = "程序已确认对应关系"
        elif risk in {RiskLevel.HIGH.value, RiskLevel.UNKNOWN.value}:
            item["核对方式"] = "人工全查"
            item["入选原因"] = item["入选原因"].rstrip("；") or (
                "核对范围受限，全部核对" if risk == RiskLevel.UNKNOWN.value else "影响金额超过实际执行重要性水平，全部核对")
        else:
            item["核对方式"] = "留存备查"
            item["入选原因"] = "影响不超过明显微小错报临界值，留存备查"

    mediums = sorted([item for item in items if item["风险等级"] == RiskLevel.MEDIUM.value],
                     key=lambda item: (item["首次日期"], item["事项编号"]))
    high_count = sum(item["风险等级"] == RiskLevel.HIGH.value for item in items)
    sample_size = min(len(mediums), high_count)
    picked = {int(index * len(mediums) / sample_size) for index in range(sample_size)} if sample_size else set()
    for index, item in enumerate(mediums):
        selected = index in picked
        item["核对方式"] = "人工抽样" if selected else "留存备查"
        item["入选原因"] = (
            f"中风险按日期等距抽样，{len(mediums)}项抽取{sample_size}项，本项抽中"
            if selected else "中风险等距抽样未抽中，本次留存备查")
    return items
