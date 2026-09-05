"""从已有文字提取业务范围；完整组直接求和，不枚举工资子集。"""
import re
import unicodedata
from collections import defaultdict

from matching_policy import _critical_field_category, _normalize_text

_ID_MARKER = r"业务(?:编?号|编码)|批次(?:编?号)?|结算(?:编?号)|订单(?:编?号)|合同(?:编?号)"
_GENERIC_SUMMARIES = {"转账", "汇款", "收款", "付款", "货款", "往来款"}


def _normalize_id(value):
    """编号仅统一全半角及大小写，保留前导零和分隔符。"""
    return unicodedata.normalize("NFKC", str(value)).strip().casefold()


def _id_category(label):
    return next(name for name in ("业务", "批次", "结算", "订单", "合同") if name in label)


def row_business(row):
    fields = dict(row.get("aux_text_fields", {}) or {})
    fields.setdefault("摘要", row.get("summary", ""))
    ids = set()
    texts = []
    parties = defaultdict(set)
    for label, value in fields.items():
        value = str(value).strip()
        if not _normalize_text(value):
            continue
        if re.search(_ID_MARKER, str(label)):
            ids.add((_id_category(str(label)), _normalize_id(value)))
        if _critical_field_category(str(label)):
            parties[_critical_field_category(str(label))].add(_normalize_text(value))
        else:
            texts.append(value)
            for match in re.finditer(r"(" + _ID_MARKER + r")\s*[:：#]?\s*([A-Za-z0-9][A-Za-z0-9_./-]*)", value):
                ids.add((_id_category(match.group(1)), _normalize_id(match.group(2))))
    text = " ".join(texts)
    salary = bool(re.search(r"工资|薪资|薪酬|薪金", text))
    periods = {(int(year) if year else 0, int(month))
               for year, month in re.findall(r"(?:(\d{4})年)?(\d{1,2})月", text)} if salary else set()
    # 只有明确的汇总称谓才与员工明细区分；实际姓名继续参与冲突检查。
    if salary:
        parties["对方户名"] -= {"工资", "工资汇总", "代发工资", "员工工资", "职工工资", "工资总额"}
    return {"ids": ids, "salary": salary, "periods": periods,
            "summary": _normalize_text(row.get("summary", "")),
            "parties": dict(parties), "date": row["date"],
            "voucher": _normalize_text(row.get("voucher_no", "")),
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
    if (any(len(values) > 1 for side in id_values for values in side.values())
            or any(id_values[0][key] != id_values[1][key] for key in id_values[0].keys() & id_values[1].keys())):
        conflicts.append("业务编号或批次不同")
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
    for category in parties[0].keys() & parties[1].keys():
        if parties[0][category] and parties[1][category] and parties[0][category] != parties[1][category]:
            conflicts.append(category)
    common_id = bool(ids[0] & ids[1])
    summaries = [{r["summary"] for r in rows if r["summary"]} for rows in (bank, journal)]
    same_summary = bool(summaries[0] and summaries[0] == summaries[1]
                        and len(summaries[0]) == 1 and not summaries[0] <= _GENERIC_SUMMARIES)
    same_party = any(len(parties[0][key]) == 1 and parties[0][key] == parties[1][key]
                     for key in parties[0].keys() & parties[1].keys())
    if common_id:
        strength, basis = 4, "共同业务编号或批次：" + "、".join(category + " " + value for category, value in sorted(ids[0] & ids[1]))
    elif all(salary_sides):
        strength, basis = 2, "工资用途及发放日期、期间对应"
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
            "business_conflicts": tuple(conflicts)}


def complete_groups(profiles, window_days):
    """明确编号优先；无编号时允许同日工资或同一摘要组。"""
    grouped = defaultdict(list)
    for index, row in profiles.items():
        if row["ids"]:
            key = ("编号", tuple(sorted(row["ids"])), row["sign"])
        elif row["salary"]:
            key = ("工资", tuple(sorted(row["periods"])), row["date"].normalize(), row["sign"])
        elif row["summary"]:
            key = ("摘要", row["summary"], row["sign"], tuple(sorted(row["parties"].get("对方户名", ()))))
        else:
            continue
        grouped[key].append(index)
    result = []
    for key, indices in grouped.items():
        chunk = []
        for index in sorted(indices, key=lambda i: (profiles[i]["date"], i)):
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
            bool(candidate.evidence.get("complete_business_id")),
            candidate.metrics.total_diff_li == 0)


def candidate_sort_key(candidate):
    priority = relationship_priority(candidate)
    complete = candidate.evidence.get("resolves_full_group") and not has_business_conflict(candidate)
    return (tuple(-int(value) for value in priority), -int(bool(complete)),
            -candidate.scores.total, candidate.metrics.total_diff_li,
            candidate.date_span_days, len(candidate.bank_idxs) + len(candidate.journal_idxs),
            candidate.bank_idxs, candidate.journal_idxs, candidate.candidate_id)
