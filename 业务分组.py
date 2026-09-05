"""从已有文字提取业务范围；完整组直接求和，不枚举工资子集。"""
import re
import unicodedata
from collections import defaultdict

from matching_policy import _critical_field_category, _normalize_text

_GROUP_ID_MARKER = (
    r"业务(?:编?号|编码)|批次(?:编?号)?|(?:银企)?批量(?:编?号)|"
    r"结算(?:编?号)|订单(?:编?号)|合同(?:编?号)"
)
_TRANSACTION_ID_MARKER = r"回单(?:编?号)?|(?:交易)?流水(?:编?号)?|交易编号"
_ID_MARKER = rf"(?:{_GROUP_ID_MARKER}|{_TRANSACTION_ID_MARKER})"
_GENERIC_SUMMARIES = {"转账", "汇款", "收款", "付款", "货款", "往来款"}
_GROUP_ID_PRIORITY = ("批次", "业务", "结算", "合同", "订单", "其他编号")


def _is_generic_summary(value):
    normalized = _normalize_text(value)
    if normalized in _GENERIC_SUMMARIES:
        return True
    return bool(re.fullmatch(r"(?:支付|付|收取|收到|采购|销售)?(?:转账|汇款|收款|付款|货款|往来款)", normalized))


def _batch_category(text):
    """只识别文字中已经明确表达的批量业务类型。"""
    normalized = _normalize_text(text)
    if re.search(r"工资|薪资|薪酬|薪金", normalized):
        return "工资"
    if re.search(r"奖金|年终奖|绩效奖", normalized):
        return "奖金"
    if "员工报销" in normalized:
        return "员工报销"
    if "报销" in normalized:
        return "报销"
    if re.search(r"社保|公积金", normalized):
        return "社保公积金"
    if "税" in normalized and re.search(r"代扣|批量|缴纳|申报", normalized):
        return "税费代扣"
    if re.search(r"批量|批付|批收", normalized):
        if re.search(r"客户|收款|回款|批收", normalized):
            return "客户批收"
        if re.search(r"供应商|付款|支付|批付", normalized):
            return "供应商批付"
        return "批量收付"
    return ""


def _normalize_id(value, numeric_integer_value=""):
    """统一编号文字；仅对已证明来自数值单元格的整数去掉无意义小数点。"""
    if value is None:
        return ""
    source = numeric_integer_value or value
    normalized = unicodedata.normalize("NFKC", str(source)).strip().casefold()
    return "" if normalized in {"", "nan", "none", "null"} else normalized


def _id_category(label):
    if "批量" in label:
        return "批次"
    for name in ("业务", "批次", "结算", "订单", "合同"):
        if name in label:
            return name
    if "回单" in label or "流水" in label or "交易编号" in label:
        return "交易流水"
    return "其他编号"


def _group_boundary_ids(ids):
    """选择当前行最高层级的组边界；批次内订单号只作为子项证据。"""
    by_category = defaultdict(set)
    for category, value in ids:
        by_category[category].add(value)
    for category in _GROUP_ID_PRIORITY:
        if by_category[category]:
            return frozenset((category, value) for value in by_category[category])
    return frozenset()


def row_business(row):
    fields = dict(row.get("aux_text_fields", {}) or {})
    field_evidence = row.get("aux_text_evidence", {})
    if not isinstance(field_evidence, dict):
        field_evidence = {}
    fields.setdefault("摘要", row.get("summary", ""))
    ids = set()
    transaction_ids = set()
    texts = []
    parties = defaultdict(set)
    for label, value in fields.items():
        value = str(value).strip()
        if not _normalize_text(value):
            continue
        if re.search(_ID_MARKER, str(label)):
            evidence = field_evidence.get(str(label), {})
            numeric_integer_value = (
                evidence.get("numeric_integer_value", "")
                if isinstance(evidence, dict)
                else ""
            )
            identifier = (
                _id_category(str(label)),
                _normalize_id(value, numeric_integer_value),
            )
            if identifier[0] == "交易流水":
                transaction_ids.add(identifier)
            else:
                ids.add(identifier)
        critical_category = _critical_field_category(str(label))
        if critical_category == "交易流水号":
            # 回单号、银行交易流水号通常每笔都不同，只能用于一对一核验，
            # 不能作为工资等批次的共同编号或组内冲突依据。
            continue
        if critical_category:
            parties[critical_category].add(_normalize_text(value))
        else:
            texts.append(value)
            for match in re.finditer(r"(" + _ID_MARKER + r")\s*[:：#]?\s*([A-Za-z0-9][A-Za-z0-9_./-]*)", value):
                identifier = (_id_category(match.group(1)), _normalize_id(match.group(2)))
                if identifier[0] == "交易流水":
                    transaction_ids.add(identifier)
                else:
                    ids.add(identifier)
    text = " ".join(texts)
    row_date = row["date"]
    batch_category = _batch_category(text)
    salary = batch_category == "工资"
    periods = set()
    if batch_category:
        for year, month in re.findall(r"(?:(\d{4})年)?(\d{1,2})月", text):
            month_number = int(month)
            if year:
                year_number = int(year)
            else:
                year_number = int(row_date.year)
                # 年初出现“12月/11月”通常指上年工资，避免跨年误并；
                # 其他月份以交易日期年度补全，解决一侧写年份、一侧省略年份。
                if int(row_date.month) <= 2 and month_number >= 11:
                    year_number -= 1
            periods.add((year_number, month_number))
    # 只有明确的汇总称谓才与员工明细区分；实际姓名继续参与冲突检查。
    if salary:
        parties["对方户名"] -= {"工资", "工资汇总", "代发工资", "员工工资", "职工工资", "工资总额"}
    return {"ids": ids, "transaction_ids": transaction_ids,
            "salary": salary, "batch_category": batch_category,
            "periods": periods,
            "summary": _normalize_text(row.get("summary", "")),
            "parties": dict(parties), "date": row_date,
            "voucher_word": _normalize_text(row.get("voucher_word", "")),
            "voucher": _normalize_id(row.get("voucher_no", "")),
            "sign": 1 if row["amount_decimal"] > 0 else -1 if row["amount_decimal"] < 0 else 0}


def business_evidence(bank, journal):
    """按原始每行检查冲突，避免把多名员工拼成一个交易对手。"""
    conflicts = []
    ids = [set().union(*(r["ids"] for r in rows)) for rows in (bank, journal)]
    id_values = []
    for values in ids:
        by_category = defaultdict(set)
        for category, value in values:
            by_category[category].add(value)
        id_values.append(by_category)
    common_identifiers = set()
    # 一个候选若单侧已经混入多个明示批次，范围本身就不完整。
    if any(len(side.get("批次", set())) > 1 for side in id_values):
        conflicts.append("业务编号或批次不同")
    boundary_category = next(
        (
            category for category in _GROUP_ID_PRIORITY
            if id_values[0].get(category) and id_values[1].get(category)
        ),
        "",
    )
    if boundary_category:
        left = id_values[0][boundary_category]
        right = id_values[1][boundary_category]
        intersection = left & right
        # 采用双方共同存在的最高层边界。该边界必须完整且唯一地相同；
        # 若共同批次已经相同，批次内订单、合同等较低层子项仍可不同。
        if not intersection or left != right or len(left) != 1:
            if "业务编号或批次不同" not in conflicts:
                conflicts.append("业务编号或批次不同")
        else:
            common_identifiers = {
                (boundary_category, value) for value in intersection
            }
    transaction_ids = [
        set().union(*(r.get("transaction_ids", set()) for r in rows))
        for rows in (bank, journal)
    ]
    aggregate_batch_context = bool(
        any(category in {"批次", "业务"} for category, _ in common_identifiers)
        or all(all(row.get("salary") for row in rows) for rows in (bank, journal))
        or (
            all(all(row.get("batch_category") for row in rows) for rows in (bank, journal))
            and {
                row.get("batch_category") for row in bank
            } == {
                row.get("batch_category") for row in journal
            }
        )
    )
    one_side_is_single = len(bank) == 1 or len(journal) == 1
    if (
        one_side_is_single
        and not aggregate_batch_context
        and all(transaction_ids)
        and transaction_ids[0] != transaction_ids[1]
    ):
        conflicts.append("交易流水号不同")
    common_transaction_id = bool(
        one_side_is_single
        and not aggregate_batch_context
        and all(transaction_ids)
        and transaction_ids[0] == transaction_ids[1]
    )
    periods = [set().union(*(r["periods"] for r in rows)) for rows in (bank, journal)]
    months = [{month for _, month in values} for values in periods]
    years = [{year for year, _ in values if year} for values in periods]
    if (any(len(values) > 1 for values in months + years)
            or (all(months) and months[0] != months[1])
            or (all(years) and years[0] != years[1])):
        conflicts.append("工资期间不同")
    salary_sides = [all(r["salary"] for r in rows) for rows in (bank, journal)]
    if any(r["salary"] for r in bank + journal) and not all(salary_sides):
        conflicts.append("工资与其他用途混合")
    batch_categories = [
        {r.get("batch_category", "") for r in rows if r.get("batch_category", "")}
        for rows in (bank, journal)
    ]
    batch_sides = [all(r.get("batch_category", "") for r in rows) for rows in (bank, journal)]
    if any(batch_categories) and not all(batch_sides):
        conflicts.append("批量业务与其他用途混合")
    elif (
        any(len(values) > 1 for values in batch_categories)
        or (
            all(batch_categories)
            and batch_categories[0] != batch_categories[1]
        )
    ):
        conflicts.append("批量业务类型不同")
    common_batch_category = (
        next(iter(batch_categories[0]))
        if all(len(values) == 1 for values in batch_categories)
        and batch_categories[0] == batch_categories[1]
        else ""
    )
    vouchers = defaultdict(set)
    for row in journal:
        if row["voucher"]:
            vouchers[row["voucher"]].add(row["date"].strftime("%Y-%m"))
    if any(len(period) > 1 for period in vouchers.values()):
        conflicts.append("同凭证号跨年度或期间")
    parties = []
    for rows in (bank, journal):
        values = defaultdict(set)
        for row in rows:
            for label, names in row["parties"].items():
                values[label].update(names)
        parties.append(values)
    aggregate_party_hierarchy = bool(
        (
            any(category == "批次" for category, _ in common_identifiers)
            or all(salary_sides)
            or bool(common_batch_category)
        )
        and ((len(bank) == 1) != (len(journal) == 1))
    )
    aggregate_party_categories = []
    for category in parties[0].keys() & parties[1].keys():
        if (
            parties[0][category]
            and parties[1][category]
            and parties[0][category] != parties[1][category]
        ):
            if aggregate_party_hierarchy:
                aggregate_party_categories.append(category)
            else:
                conflicts.append(category)
    if aggregate_party_categories:
        aggregate_party_hierarchy_hint = (
            "银行流水汇总主体与序时账明细主体层级不同，需结合批量业务清单核查"
            if len(bank) == 1
            else "银行流水明细主体与序时账汇总主体层级不同，需结合批量业务清单核查"
        )
    else:
        aggregate_party_hierarchy_hint = ""
    common_id = bool(common_identifiers)
    summaries = [{r["summary"] for r in rows if r["summary"]} for rows in (bank, journal)]
    same_summary = bool(
        summaries[0]
        and summaries[0] == summaries[1]
        and len(summaries[0]) == 1
        and not _is_generic_summary(next(iter(summaries[0])))
    )
    same_party = any(len(parties[0][key]) == 1 and parties[0][key] == parties[1][key]
                     for key in parties[0].keys() & parties[1].keys())
    if common_id:
        strength, basis = 4, "共同业务编号或批次：" + "、".join(
            category + " " + value
            for category, value in sorted(common_identifiers)
        )
    elif common_transaction_id:
        strength, basis = 4, "共同交易流水号或回单号"
    elif all(salary_sides):
        strength, basis = 2, "工资用途及发放日期、期间对应"
    elif common_batch_category:
        strength, basis = 2, f"{common_batch_category}用途及业务日期、期间对应"
    elif same_party and same_summary:
        strength, basis = 3, "明确对方及具体摘要共同对应"
    elif same_party:
        strength, basis = 2, "明确对方信息对应"
    elif same_summary:
        strength, basis = 1, "摘要对应"
    else:
        strength, basis = 0, "按金额、日期及文字相似程度形成候选"
    return {"business_strength": strength, "business_basis": basis,
            "shared_business_id": common_id,
            "shared_business_category": boundary_category if common_id else "",
            "shared_business_identifiers": tuple(sorted(common_identifiers)),
            "shared_transaction_id": common_transaction_id,
            "salary_group": all(salary_sides),
            "batch_category": common_batch_category,
            "specific_summary_group": same_summary,
            "same_party_group": same_party,
            "aggregate_party_hierarchy_hint": aggregate_party_hierarchy_hint,
            "aggregate_party_categories": tuple(sorted(aggregate_party_categories)),
            "business_conflicts": tuple(conflicts)}


def complete_groups(profiles, window_days, source_type=""):
    """日记账完整凭证优先；其余按编号、批量用途或具体摘要组成完整范围。"""
    grouped = defaultdict(list)
    for index, row in profiles.items():
        if source_type == "journal" and row["voucher"]:
            key = (
                "凭证",
                row["date"].strftime("%Y-%m"),
                row.get("voucher_word", ""),
                row["voucher"],
                row["sign"],
            )
        elif row["ids"]:
            key = ("编号", tuple(sorted(_group_boundary_ids(row["ids"]))), row["sign"])
        elif row["voucher"]:
            key = (
                "凭证",
                row["date"].strftime("%Y-%m"),
                row.get("voucher_word", ""),
                row["voucher"],
                row["sign"],
            )
        elif row["salary"]:
            key = ("工资", tuple(sorted(row["periods"])), row["date"].normalize(), row["sign"])
        elif row.get("batch_category"):
            key = (
                "批量业务",
                row["batch_category"],
                tuple(sorted(row["periods"])),
                row["date"].normalize(),
                row["sign"],
            )
        elif row["summary"] and not _is_generic_summary(row["summary"]):
            key = ("摘要", row["summary"], row["sign"], tuple(sorted(row["parties"].get("对方户名", ()))))
        else:
            continue
        grouped[key].append(index)
    result = []
    for key, indices in grouped.items():
        ordered = sorted(indices, key=lambda i: (profiles[i]["date"], i))
        if key[0] == "凭证":
            # 同一完整年月、凭证号和方向是一个原子组；超窗口时由跨侧检查整组拒绝。
            result.append((key, tuple(ordered)))
            continue
        chunk = []
        for index in ordered:
            if chunk and (profiles[index]["date"] - profiles[chunk[0]]["date"]).days > window_days:
                result.append((key, tuple(chunk)))
                chunk = []
            chunk.append(index)
        if chunk:
            result.append((key, tuple(chunk)))
    return result


def has_business_conflict(candidate):
    return bool(candidate.evidence.get("business_conflicts") or
                (candidate.text_evidence and candidate.text_evidence.conflicting_fields))


def relationship_priority(candidate):
    """先业务依据，再完整无差额；分数只在同一业务层次内排序。"""
    strength = candidate.evidence.get("business_strength", 0)
    return (not has_business_conflict(candidate),
            strength if strength >= 2 else 0,
            bool(candidate.evidence.get("complete_business_id")
                 or candidate.evidence.get("complete_business_group")),
            bool(candidate.evidence.get("represents_full_observed_group")),
            candidate.metrics.total_diff_li == 0)


def candidate_sort_key(candidate):
    priority = relationship_priority(candidate)
    complete = candidate.evidence.get("resolves_full_group") and not has_business_conflict(candidate)
    observed_group = candidate.evidence.get("represents_full_observed_group", False)
    return (tuple(-int(value) for value in priority), -int(bool(complete)),
            -int(bool(observed_group)),
            -candidate.scores.total, candidate.metrics.total_diff_li,
            candidate.date_span_days, len(candidate.bank_idxs) + len(candidate.journal_idxs),
            candidate.match_type,
            getattr(candidate, "stable_key", ""),
            getattr(candidate, "composition_key", ""),
            candidate.bank_idxs, candidate.journal_idxs, candidate.candidate_id)
