"""按工作表读取原件，使用已有结构识别和 Agent 的列映射。"""

from dataclasses import dataclass
from pathlib import Path
import re
from zipfile import BadZipFile

import pandas as pd

from input_precheck import TableStructure, detect_table_structure, flatten_header_rows
from utils import clean_amount, parse_date
from 列识别 import ColumnRecognitionError, FIELD_LABELS, normalized_identity, recognize_columns
from 核对期间 import add_balance_point


@dataclass
class 原表:
    raw: pd.DataFrame
    structure: TableStructure
    mapping: dict
    facts: dict


def _facts(prefix, origin):
    dates = r"(\d{4}[-/.年]\d{1,2}[-/.月]\d{1,2}日?)"
    lines = [" ".join(str(v) for v in row if v is not None) for row in prefix]
    facts = {"balances": {}, "periods": []}
    periods = facts["periods"]
    for line in lines:
        if re.search(r"查询|账单期间|对账期间|起止|起讫|统计区间|记账期间", line):
            found = [parse_date(value) for value in re.findall(dates, line)]
            if len(found) >= 2 and all(day is not None for day in found[:2]) and found[0] <= found[1]:
                periods.append((found[0].normalize(), found[1].normalize(), "完整", origin))
    if periods:
        text = " ".join(lines)
        zero = re.search(r"(?:交易|记录|流水)?\s*(?:笔数\s*[:：]?\s*0(?!\d)|0\s*笔)|无交易|无发生", text)
        incoming = re.search(r"(?:收入|贷方发生额)\s*[:：]?\s*0(?:\.0+)?(?![\d.])", text)
        outgoing = re.search(r"(?:支出|借方发生额)\s*[:：]?\s*0(?:\.0+)?(?![\d.])", text)
        if zero and incoming and outgoing:
            facts["periods"] = [(start, end, "无交易", source) for start, end, _, source in periods]
        for line in lines:
            point = re.search(dates + r"\s*(?:日末|期末|月末)?余额\s*[:：]?\s*([+-]?[\d,]+(?:\.\d+)?)", line)
            if point:
                day = parse_date(point[1])
                amount = clean_amount(point[2])
                if day is not None and amount is not None:
                    add_balance_point(facts, day.normalize(), (amount, origin))
    return facts


def _pages(path, loader):
    if path.suffix.lower() == ".xlsx":
        from 直读XML import 读簿
        sheets = list(读簿(path, 数据=True))
    else:
        sheets = [None]
    for sheet in sheets:
        try:
            structure = detect_table_structure(path, sheet=sheet)
        except ValueError:
            if sheet is None:
                raise
            structure = None
        prefix = [] if sheet is None else list(sheet.行值(1, min(sheet.最大行, 40)))
        origin = f"{path.name} · {sheet.名称}" if sheet is not None else path.name
        facts = _facts(prefix, origin)
        no_activity = any(p[2] == "无交易" for p in facts["periods"])
        if no_activity and sheet is not None:
            for index, row in enumerate(prefix):
                columns = flatten_header_rows([row])
                empty = pd.DataFrame(columns=columns)
                if not recognize_columns(empty, "bank").problems:
                    structure = TableStructure(index, 1, columns, 0)
                    break
        if structure is None:
            continue
        if sheet is not None:
            facts = _facts(prefix[:structure.skiprows + structure.header_rows], origin)
        raw = loader.load_file(path, structure.skiprows, structure.header_rows, structure.columns, sheet=sheet)
        yield raw, structure, facts, origin


def file_identity(path, side):
    """文件名仅用于补全公司或开户行，完整账号优先从原表取值。"""
    from data_loader import DataLoader
    from input_precheck import statement_filename_identity

    named = statement_filename_identity(path)
    labels = (("本方公司", "本方公司名称", "公司名称", "企业名称", "本方户名"),
              ("本方开户行", "开户银行", "开户行", "银行名称"))
    signals = [set(), set(), set()]
    try:
        for raw, _, _, _ in _pages(path, DataLoader()):
            account = recognize_columns(raw, side).mapping.get("account")
            for index, columns in enumerate((*labels, (account,))):
                for column in columns:
                    if column in raw:
                        signals[index].update(normalized_identity(v, "account") if index == 2 else str(v).strip()
                                              for v in raw[column].dropna() if str(v).strip())
    except (OSError, ValueError, KeyError, BadZipFile):
        return named
    if any(len(values) > 1 for values in signals):
        return None
    identity = tuple(next(iter(values), named[index] if named else "") for index, values in enumerate(signals))
    return identity if identity[-1] else named


def read_source(path, side, loader, *, overrides=None, review=None):
    path = Path(path)
    pages = []
    combined_facts = {"periods": [], "balances": {}}
    for raw, structure, facts, origin in _pages(path, loader):
        columns = dict(overrides or {})
        roles = recognize_columns(raw, side, overrides=columns)
        if not roles.mapping.get("date") and any(re.search(r"合计|累计|汇总|笔数|统计", col) for col in raw):
            # 有发生额合计和笔数、没有交易日期的统计表不当作续页。
            continue
        used = {str(value) for value in roles.mapping.values() if isinstance(value, str)}
        used.update(roles.mapping.get("auxiliary_text_columns", ()))
        unclassified = [col for col in raw if col not in used and not col.startswith("__")
                        and re.search(r"号|账号|日期|时间|用途|参考|付款|收款|金额|amount|date|ref\b|number", col, re.I)
                        and not re.search(r"可用|起息|手续费|打印|合计|小计|振替|余额", col)]
        if (roles.choices or roles.problems or unclassified) and review:
            packet = {"kind": "columns", "side": side, "origin": origin,
                      "fields": list(FIELD_LABELS), "choices": roles.choices, "problems": roles.problems,
                      "unclassified": unclassified,
                      "preview": raw.head(8).fillna("").astype(str).to_dict("records")}
            for _ in range(2 if roles.choices or roles.problems else 1):
                answer, _ = review(packet)
                mapping = (answer or {}).get("columns")
                supplied = mapping.get(side) if isinstance(mapping, dict) else None
                if not isinstance(supplied, dict) or not all(k in FIELD_LABELS and isinstance(v, str) for k, v in supplied.items()):
                    continue
                updated = recognize_columns(raw, side, overrides={**columns, **supplied})
                if not updated.problems and not updated.choices:
                    roles = updated
                    break
        if roles.choices or roles.problems:
            raise ColumnRecognitionError([{"侧": side, "文件": path.name, "工作表": origin,
                                           "候选列": roles.choices, "问题": roles.problems}])
        mapping = roles.mapping
        for start, end, kind, source in facts["periods"]:
            if kind == "无交易" and not raw[mapping["date"]].dropna().empty:
                continue
            combined_facts["periods"].append((start, end, kind, source))
        for day, value in facts["balances"].items():
            add_balance_point(combined_facts, day, value)
        for day, points in facts.get("balance_conflicts", {}).items():
            for value in points:
                add_balance_point(combined_facts, day, value)
        pages.append((raw, structure, mapping))
    if not pages:
        raise ValueError(f"原表没有找到交易明细页：{path.name}")
    first, structure, mapping = pages[0]
    seen = []
    distinct = []
    for raw, page_info, own in pages:
        fields = [column for column in raw if not str(column).startswith("__")]
        original = raw[fields]
        reliable_ids = [column for column in fields if re.search(r"(?:交易|银行)?流水(?:编?号)|回单(?:编?号)", column)]
        exact_copy = any(original.equals(previous) for previous in seen)
        if raw.empty or exact_copy and any(raw[column].notna().all() and raw[column].is_unique for column in reliable_ids):
            continue
        seen.append(original)
        distinct.append((raw, page_info, own))
    if any(own["mode"] != mapping["mode"] for _, _, own in distinct):
        return _combine_layouts(loader, side, distinct, structure, combined_facts)
    sections = []
    for raw, _, own in distinct:
        renamed = {}
        for role in ("date", "amount", "debit", "credit", "direction", "balance", "summary", "account", "currency",
                     "voucher", "voucher_word", "counterparty_account", "order", "transaction_id"):
            old = own.get(role)
            if old:
                if not mapping.get(role):
                    mapping[role] = old
                renamed[old] = mapping[role]
        mapping["auxiliary_text_columns"] = list(dict.fromkeys(
            [*mapping.get("auxiliary_text_columns", ())] +
            [renamed.get(column, column) for column in own.get("auxiliary_text_columns", ())]
        ))
        renamed = {old: new for old, new in renamed.items() if old != new}
        if renamed:
            raw = raw.drop(columns=[column for column in renamed.values() if column in raw and column not in renamed]).rename(columns=renamed)
        sections.append(raw)
    result = pd.concat(sections, ignore_index=True) if sections else first
    return 原表(result, structure, mapping, combined_facts)


def _combine_layouts(loader, side, pages, structure, facts):
    sections = []
    auxiliary = set()
    roles = {"date": "交易日期", "income": "收入金额", "expense": "支出金额", "account": "本方账号",
             "currency": "币种", "balance": "交易后余额", "summary": "摘要", "voucher_no": "凭证号",
             "voucher_word": "凭证字", "original_file_row": "__file_row__", "original_sheet": "__sheet__"}
    for raw, _, mapping in pages:
        frame = loader.standardize_data(raw, mapping, side, recover_amounts=True)
        normalized = frame.rename(columns=roles)
        normalized = normalized[[label for label in roles.values() if label in normalized]].copy()
        normalized["收入金额"] = frame["income"] if "income" in frame else frame["amount"].clip(lower=0)
        normalized["支出金额"] = frame["expense"] if "expense" in frame else -frame["amount"].clip(upper=0)
        for column in {key for fields in frame["aux_text_fields"] for key in fields} - set(normalized):
            normalized[column] = frame["aux_text_fields"].map(lambda fields: fields.get(column))
            auxiliary.add(column)
        sections.append(normalized)
    combined = pd.concat(sections, ignore_index=True)
    identified = recognize_columns(combined, side).mapping
    identified["auxiliary_text_columns"] = list(dict.fromkeys([*identified["auxiliary_text_columns"], *sorted(auxiliary)]))
    return 原表(combined, structure, identified, facts)
