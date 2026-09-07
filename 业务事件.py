"""特殊银行业务识别：同侧退款链、手续费净额和重复线索。"""

from __future__ import annotations

import hashlib
from collections import defaultdict
from typing import Any

import pandas as pd

from data_structures import BusinessClue, BusinessEvent


RETURN_WORDS = ("退款", "退回", "退汇", "冲销", "撤销", "退票", "红字")
FEE_WORDS = ("手续费", "佣金", "结算费", "账户管理费", "网银费", "汇费", "转账费", "短信费")


def _text(row: pd.Series) -> str:
    fields = row.get("aux_text_fields", {})
    values = [row.get("summary", "")]
    if isinstance(fields, dict):
        values.extend(fields.values())
    return "|".join(str(value).strip() for value in values if value is not None)


def _field(row: pd.Series, names: tuple[str, ...]) -> str:
    fields = row.get("aux_text_fields", {})
    candidates: list[Any] = []
    if isinstance(fields, dict):
        for key, value in fields.items():
            if any(name in str(key) for name in names):
                candidates.append(value)
    for name in names:
        if name in row.index:
            candidates.append(row.get(name))
    for value in candidates:
        text = str(value or "").strip()
        if text and text.lower() not in {"nan", "none"}:
            return text
    return ""


def business_identity(row: pd.Series) -> tuple[str, str]:
    return (
        _field(row, ("业务编号", "批次号", "回单号", "申请单号", "订单号")),
        _field(row, ("对方", "户名", "客户", "供应商")),
    )


def has_fee_evidence(row: pd.Series) -> bool:
    text = _text(row)
    return any(word in text for word in FEE_WORDS)


def rows_share_business(left: pd.Series, right: pd.Series) -> tuple[bool, str]:
    left_id, left_party = business_identity(left)
    right_id, right_party = business_identity(right)
    if left_id and right_id and left_id == right_id:
        return True, f"共同业务编号：{left_id}"
    if left_party and right_party and left_party == right_party:
        return True, f"共同对方：{left_party}"
    return False, ""


def _refund_business(left: pd.Series, right: pd.Series) -> tuple[bool, str]:
    """退款链不能用共同对方掩盖明确编号冲突，逐笔回单号不同不视为冲突。"""
    left_party, right_party = business_identity(left)[1], business_identity(right)[1]
    if left_party and right_party and left_party != right_party:
        return False, ""
    for names in (("批次号", "批次编号"), ("业务编号",), ("订单号",), ("申请单号",)):
        left_id, right_id = _field(left, names), _field(right, names)
        if left_id and right_id:
            if left_id != right_id:
                return False, ""
            break
    return rows_share_business(left, right)


def _stable_id(prefix: str, source: str, indexes: tuple[int, ...]) -> str:
    raw = f"{prefix}|{source}|{','.join(map(str, indexes))}".encode("utf-8")
    return f"{prefix}-{hashlib.sha256(raw).hexdigest()[:12].upper()}"


def detect_same_side_events(
    frame: pd.DataFrame,
    source: str,
    date_window_days: int = 31,
) -> tuple[list[BusinessEvent], list[BusinessClue]]:
    """识别强证据同侧业务链；弱证据只生成线索，不改变原始行。"""
    if frame.empty:
        return [], []
    ordered = sorted(
        (int(index) for index in frame.index),
        key=lambda index: (pd.Timestamp(frame.loc[index, "date"]), index),
    )
    order_position = {index: position for position, index in enumerate(ordered)}
    events: list[BusinessEvent] = []
    clues: list[BusinessClue] = []
    consumed: set[int] = set()

    # 三段链：原付款/收款、明确退回、再次付款/收款。
    for middle_pos in range(1, len(ordered) - 1):
        middle_idx = ordered[middle_pos]
        if middle_idx in consumed:
            continue
        middle = frame.loc[middle_idx]
        if not any(word in _text(middle) for word in RETURN_WORDS):
            continue
        for left_idx in reversed(ordered[:middle_pos]):
            if left_idx in consumed:
                continue
            left = frame.loc[left_idx]
            if abs(int(left["amount_decimal"])) != abs(int(middle["amount_decimal"])):
                continue
            if int(left["amount_decimal"]) * int(middle["amount_decimal"]) >= 0:
                continue
            if (pd.Timestamp(middle["date"]) - pd.Timestamp(left["date"])).days > date_window_days:
                continue
            shared_left, basis_left = _refund_business(left, middle)
            if not shared_left:
                continue
            for right_idx in ordered[middle_pos + 1:]:
                if right_idx in consumed:
                    continue
                right = frame.loc[right_idx]
                if int(right["amount_decimal"]) != int(left["amount_decimal"]):
                    continue
                if (pd.Timestamp(right["date"]) - pd.Timestamp(middle["date"])).days > date_window_days:
                    continue
                shared_right, basis_right = _refund_business(middle, right)
                if not shared_right:
                    continue
                if not _refund_business(left, right)[0]:
                    continue
                indexes = (left_idx, middle_idx, right_idx)
                dates = [pd.Timestamp(frame.loc[index, "date"]) for index in indexes]
                payment = int(left["amount_decimal"]) < 0
                event_type = "退汇后重付" if payment else "退款后重新收款"
                formula = "原业务＋退回＋重付＝最终净影响" if payment else "原收款＋退款＋重新收款＝最终净影响"
                events.append(
                    BusinessEvent(
                        event_id=_stable_id("EVT", source, indexes),
                        source=source,
                        event_type=event_type,
                        row_idxs=indexes,
                        net_amount_li=sum(int(frame.loc[index, "amount_decimal"]) for index in indexes),
                        relationship_formula=formula,
                        evidence_basis=f"{basis_left}；{basis_right}；退回文字和日期先后关系",
                        is_cross_period=len({(date.year, date.month) for date in dates}) > 1,
                    )
                )
                consumed.update(indexes)
                break
            if middle_idx in consumed:
                break

    # 两段链：原业务与明确退款/退汇或红字冲销。
    for right_idx in ordered:
        if right_idx in consumed:
            continue
        right = frame.loc[right_idx]
        if not any(word in _text(right) for word in RETURN_WORDS):
            continue
        for left_idx in reversed(ordered[:order_position[right_idx]]):
            if left_idx in consumed:
                continue
            left = frame.loc[left_idx]
            if not int(left["amount_decimal"]) or int(left["amount_decimal"]) != -int(right["amount_decimal"]):
                continue
            if (pd.Timestamp(right["date"]) - pd.Timestamp(left["date"])).days > date_window_days:
                continue
            shared, basis = _refund_business(left, right)
            if not shared:
                continue
            indexes = (left_idx, right_idx)
            dates = [pd.Timestamp(left["date"]), pd.Timestamp(right["date"])]
            event_type = "付款退回" if int(left["amount_decimal"]) < 0 else "收款退款"
            if "冲销" in _text(right) or "红字" in _text(right):
                event_type = "账内冲销"
            events.append(
                BusinessEvent(
                    event_id=_stable_id("EVT", source, indexes),
                    source=source,
                    event_type=event_type,
                    row_idxs=indexes,
                    net_amount_li=0,
                    relationship_formula="原业务＋退款或冲销＝零",
                    evidence_basis=f"{basis}；明确退款、退汇或冲销文字；日期先后关系",
                    is_cross_period=len({(date.year, date.month) for date in dates}) > 1,
                )
            )
            consumed.update(indexes)
            break

    # 完全重复和同额重复线索。
    duplicate_groups: dict[tuple[Any, ...], list[int]] = defaultdict(list)
    for index in ordered:
        row = frame.loc[index]
        txid = _field(row, ("流水号", "交易号", "回单号"))
        identity = business_identity(row)
        key = (pd.Timestamp(row["date"]), int(row["amount_decimal"]), str(row.get("summary", "")).strip(), identity)
        duplicate_groups[key].append(index)
    for key, indexes in duplicate_groups.items():
        if len(indexes) < 2:
            continue
        txids = [_field(frame.loc[index], ("流水号", "交易号", "回单号")) for index in indexes]
        same_nonempty_txid = bool(txids[0]) and len(set(txids)) == 1
        clue_type = "强重复线索" if same_nonempty_txid else "同额重复疑点"
        rows = tuple(indexes)
        clues.append(
            BusinessClue(
                clue_id=_stable_id("DUP", source, rows),
                source=source,
                clue_type=clue_type,
                row_idxs=rows,
                reason=("流水号相同且日期、金额、对方和摘要一致，保留全部原行" if same_nonempty_txid else "日期、金额、对方和摘要一致但流水号不同或缺失，不能自动删除"),
                impact_amount_li=sum(abs(int(frame.loc[index, "amount_decimal"])) for index in rows),
            )
        )

    # 未进入强业务链的正负同额记录，只列线索。
    clue_pairs: set[tuple[int, int]] = set()
    indexes_by_amount: dict[int, list[int]] = defaultdict(list)
    for index in ordered:
        indexes_by_amount[int(frame.loc[index, "amount_decimal"])].append(index)
    for left_idx in ordered:
        if left_idx in consumed:
            continue
        left = frame.loc[left_idx]
        if not int(left["amount_decimal"]):
            continue
        opposite_indexes = indexes_by_amount.get(-int(left["amount_decimal"]), [])
        for right_idx in opposite_indexes:
            if order_position[right_idx] <= order_position[left_idx]:
                continue
            if right_idx in consumed:
                continue
            right = frame.loc[right_idx]
            days = abs((pd.Timestamp(right["date"]) - pd.Timestamp(left["date"])).days)
            if days > date_window_days:
                continue
            pair = (left_idx, right_idx)
            if pair in clue_pairs:
                continue
            shared, basis = rows_share_business(left, right)
            left_party = business_identity(left)[1]
            right_party = business_identity(right)[1]
            reason = "只有正负同额，缺少共同业务编号或对方证据"
            if left_party and right_party and left_party != right_party:
                reason = "正负同额但对方不一致，不能认定为退款或冲销"
            elif shared:
                reason = f"{basis}，但缺少完整的原业务、退回或重付链"
            clues.append(
                BusinessClue(
                    clue_id=_stable_id("OPP", source, pair),
                    source=source,
                    clue_type="正负同额线索",
                    row_idxs=pair,
                    reason=reason,
                    impact_amount_li=abs(int(left["amount_decimal"])),
                )
            )
            clue_pairs.add(pair)
            break
    return events, clues
