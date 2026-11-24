"""根据原表字段含义和实际行值确定交易列。"""

from dataclasses import dataclass
from itertools import product
import re
import unicodedata

import pandas as pd

from matching_policy import COUNTERPARTY_FIELD_KEYWORDS
from utils import parse_date


FIELD_LABELS = {
    "date": "日期", "amount": "单列金额", "debit": "借方/支出金额", "credit": "贷方/收入金额",
    "direction": "收支方向", "balance": "交易后余额", "account": "本方账号", "currency": "币种",
    "summary": "摘要", "voucher": "凭证号", "voucher_word": "凭证字", "mode": "金额列格式",
    "counterparty_account": "对方账号", "order": "订单编号", "transaction_id": "交易流水号",
}
_AUXILIARY = (
    "摘要", "业务说明", "交易用途", *COUNTERPARTY_FIELD_KEYWORDS, "附言", "备注",
    "业务号", "业务编号", "批次号", "批次编号", "银企批量号", "银企批次号",
    "订单号", "订单编号", "合同号", "结算号", "流水号", "交易流水号", "回单号",
)
_OTHER_PARTY = r"对方|(?:付|收|汇)款(?:人|方|单位|户|账号|帐号|账户|卡号|币种)|counterparty|payer|payee|beneficiary|recipient|remitter"
_OTHER_AMOUNT = r"手续费|工本费|佣金|汇率|税额|利息金额|可用|冻结|余额|balance|合计|累计|笔数|exchange rate|commission|fee|(?:^| )tax"
_OTHER_DATE = r"起息|计息|结息|价值日|到期|清算日期|入单日期|打印日期|value date|interest date"


def _parts(label):
    text = unicodedata.normalize("NFKC", str(label)).strip().lower()
    parts = [re.sub(r"(?:_\d+|\.\d+)$", "", part.strip()) for part in text.split("|")]
    parts = [re.sub(r"[_\s]+", " ", part).strip() for part in parts]
    return [re.sub(r"\s*\((?:人民币|元|rmb|cny)\)$", "", part).strip() for part in parts]


def amount_basis_label(label):
    text = unicodedata.normalize("NFKC", str(label)).lower()
    original = re.search(r"原币|外币|交易币|original currency|transaction currency", text)
    base = re.search(r"本位币|本币|base currency|functional currency", text)
    return "冲突" if original and base else "原币" if original else "本位币" if base else ""


def _without_basis(text):
    return re.sub(r"原币|外币|交易币|记账本位币|本位币|本币|original currency|transaction currency|base currency|functional currency", "", text).strip()


def non_own_account_column(label):
    return bool(re.search(_OTHER_PARTY, " ".join(_parts(label))))


def direction_column(label):
    parts = _parts(label)
    return any(re.fullmatch(r"方向|(?:借贷|收支|收付|交易|进出账)(?:方向|标[志记识]|类别|类型)|借贷|收支|direction|dr/cr|cr/dr|debit/credit", part)
               for part in parts)


def _money_names(parts):
    labels = [_without_basis(part) for part in parts]
    if len(labels) > 1:
        labels.append("".join(labels[-2:]))
    return labels


def split_amount_role(label, source):
    parts = _parts(label)
    if direction_column(label) or re.search(rf"{_OTHER_AMOUNT}|账号|帐号|账户|户名|币种|标志|标记|标识|开户行", " ".join(parts)):
        return None
    debit = r"借方|借|debit"
    credit = r"贷方|贷|credit"
    outgoing = r"|支出|付款|付方|出账|取款|withdrawal"
    incoming = r"|收入|收款|收方|入账|存款|deposit"
    debit += outgoing if source == "bank" else incoming
    credit += incoming if source == "bank" else outgoing
    roles = [role for role, aliases in (("debit", debit), ("credit", credit))
             if any(re.fullmatch(rf"(?:{aliases})\s*(?:金额|发生额|发生金额|额|amount)?", name)
                    for name in _money_names(parts))]
    return roles[0] if len(roles) == 1 else None


def field_problem(role, label, source):
    text = " ".join(_parts(label))
    if role in ("account", "currency") and (non_own_account_column(label) or re.search(r"户名|账户名|开户行|account name", text)):
        return "属于其他交易字段"
    if role == "account" and re.search(r"余额|balance", text):
        return "是余额字段"
    if role == "date" and re.search(_OTHER_DATE, text):
        return "是起息、清算或其他用途的日期"
    if role == "balance" and re.search(r"可用|冻结|授信|期初|期末|available|opening|closing", text):
        return "是可用、期初期末或其他用途的余额"
    if role in ("amount", "debit", "credit"):
        if direction_column(label):
            return "是方向字段"
        split_role = split_amount_role(label, source)
        if role != "amount" and split_role and split_role != role:
            return "与所选金额方向相反"
        if re.search(rf"{_OTHER_AMOUNT}|账号|帐号|户名|币种", text):
            return "是手续费、余额或其他用途的字段"
    return ""


def _rank(role, label, source, basis=""):
    if field_problem(role, label, source):
        return 0
    parts = _parts(label)
    if role == "direction":
        return int(direction_column(label))
    if role in ("debit", "credit"):
        return int(split_amount_role(label, source) == role)
    if role == "amount":
        groups = ("交易金额|交易发生额|transaction amount", "发生额|发生金额|金额|amount|净发生额")
        parts = _money_names(parts)
    elif role == "date":
        preferred = "交易日期|交易日|transaction date" if source == "bank" else "凭证日期|凭证日|voucher date"
        timestamp = "交易时间|交易日期时间|transaction time|transaction datetime" if source == "bank" else "凭证时间|voucher datetime"
        secondary = "记账日期|入账日期|银行记账日期|posting date|booking date"
        date_names = "日期|date|datetime|时间|time"
        groups = (preferred, timestamp, secondary, date_names) if source == "bank" else (
            preferred, secondary, "交易日期|交易日|业务日期|transaction date", date_names, timestamp,
        )
    elif role == "balance":
        preferred = "交易后余额|交易余额|交易后账面余额|running balance" if source == "bank" else "账面余额|账上余额|日记账余额|ledger balance"
        groups = (preferred, "账户余额|账面余额|交易后余额|account balance", "余额|balance|结存余额|结余")
        parts = [_without_basis(part) for part in parts]
    elif role == "account":
        groups = ("本方账号|本方帐号|本户账号|本账号|本方账户|本户账户|本方卡号|own account(?: number| no)?",
                  "银行账号|银行帐号|银行账户|账户账号|账户号|账 户|账号|帐号|账户|卡号|account(?: number| no)?")
    elif role == "currency":
        tag = amount_basis_label(label)
        if tag == "冲突" or (tag and basis and tag != basis):
            return 0
        if not any(re.fullmatch(r"(?:本方|本户|交易)?(?:币种|币别|币种代码)|currency(?: code)?", _without_basis(part)) or
                   part in ("本位币", "本币", "原币", "原币代码", "本位币代码",
                            "transaction currency", "original currency", "base currency", "functional currency")
                   for part in parts):
            return 0
        return 2 if basis and tag == basis else 1
    else:
        groups = {
            "summary": ("摘要|summary", "业务说明|交易用途|交易摘要|用途", "附言|备注|说明"),
            "voucher": ("完整凭证号", "凭证号码|凭证号|凭证编号|voucher no|voucher number", "凭证|voucher"),
            "voucher_word": ("凭证字|凭证类型|凭证类别|voucher type",),
        }[role]
    return max((len(groups) - index for index, names in enumerate(groups)
                if any(re.fullmatch(names, part) for part in parts)), default=0)


def auto_select_auxiliary_columns(columns):
    return [str(column) for column in columns
            if ((non_own_account_column(column) and re.search(r"账号|帐号|账户|卡号|account", str(column).lower()))
                or (any(keyword in str(column) for keyword in _AUXILIARY)
                    and not any(keyword in str(column) for keyword in ("账号", "帐号", "卡号", "凭证"))))]


def _text(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return unicodedata.normalize("NFKC", str(value)).strip()


def normalized_identity(value, role):
    text = _text(value).upper()
    if role == "account":
        compact = re.sub(r"[\s_-]", "", text)
        return re.sub(r"^(\d+)\.0+$", r"\1", compact)
    return {"人民币": "CNY", "人民币元": "CNY", "RMB": "CNY", "156": "CNY"}.get(text, text)


def identity_columns(frame, mapping, role):
    explicit = mapping.get(role)
    if explicit in frame.columns and not field_problem(role, explicit, "bank"):
        return [explicit]
    roles = ("debit", "credit") if mapping.get("mode", "debit_credit") == "debit_credit" else ("amount",)
    tags = {amount_basis_label(mapping.get(key)) for key in roles} - {""}
    basis = mapping.get("amount_basis") or (next(iter(tags)) if len(tags) == 1 else "")
    ranked = {column: _rank(role, column, "bank", basis) for column in frame.columns}
    best = max(ranked.values(), default=0)
    return sorted(column for column, rank in ranked.items() if rank == best and rank)


@dataclass
class ColumnRecognition:
    mapping: dict
    choices: dict
    problems: list


class ColumnRecognitionError(ValueError):
    def __init__(self, issues):
        self.issues = issues
        messages = []
        for issue in issues:
            descriptions = issue["问题"] + [
                f"{FIELD_LABELS[field]}有多个候选（{'、'.join(columns)}）"
                for field, columns in issue["候选列"].items()
            ]
            messages.append(f"{issue['来源']}《{issue['文件']}》：" + "；".join(descriptions))
        super().__init__("；".join(messages))


def recognize_columns(frame, source, *, overrides=None, date_format="auto"):
    """在一张原表上选择字段。可以指定已确认的个别字段。"""
    # 延后导入。读取和输入检查也使用本模块的字段规则。
    from data_loader import directed_amount, direction_sign, parse_source_amount

    overrides = dict(overrides or {})
    requested = set(overrides)
    forced_mode = overrides.pop("mode", None)
    problems = []
    choices = {}
    mapping = {"auxiliary_text_columns": auto_select_auxiliary_columns(frame.columns)}
    for role in ("counterparty_account", "order", "transaction_id"):
        if role in overrides:
            column = overrides.pop(role)
            if column not in frame or (role == "counterparty_account" and _rank("account", column, source)):
                problems.append(f"{FIELD_LABELS[role]}列「{column}」与原表不符")
            else:
                mapping[role] = column
    series_cache = {}
    active_rows = None

    def values(role, column):
        key = (role, column)
        if key not in series_cache:
            cells = frame[column] if active_rows is None or role == "date" else frame.loc[active_rows, column]
            if role == "date":
                parser = lambda value: parse_date(value, date_format)
            elif role in ("debit", "credit", "amount", "balance"):
                parser = lambda value: parse_source_amount(value, source)
            elif role == "direction":
                parser = lambda value: direction_sign(value, source)
            elif role in ("account", "currency"):
                parser = lambda value: normalized_identity(value, role)
            else:
                parser = _text
            valid = tuple(parser(value) if _text(value) else None for value in cells)
            present = tuple(bool(_text(value)) for value in cells)
            series_cache[key] = valid, present
        return series_cache[key]

    def candidates(role, basis=""):
        if role in overrides:
            column = overrides.pop(role)
            problem = field_problem(role, column, source)
            if column not in frame.columns or problem:
                problems.append(f"{FIELD_LABELS[role]}列「{column}」" + (problem or "不在原表中"))
                return []
            return [column]

        ranked = {column: _rank(role, column, source, basis) for column in frame.columns}
        if role == "date" and not frame.empty:
            with_date = {column: rank for column, rank in ranked.items()
                         if rank and any(value is not None for value in values(role, column)[0])}
            ranked = with_date or ranked
        best = max(ranked.values(), default=0)
        names = sorted(column for column, rank in ranked.items() if rank == best and rank)
        if len(names) < 2:
            return names
        distinct = {}
        for column in names:
            tag = amount_basis_label(column) if role in ("debit", "credit", "amount", "balance", "currency") else ""
            signature = (tag, values(role, column))
            distinct.setdefault(signature, column)
        names = list(distinct.values())
        # 日期或数值填得更全，且已有有效值完全相同，才视为更完整的一列。
        return [column for column in names if not any(
            amount_basis_label(other) == amount_basis_label(column)
            and _more_complete(values(role, other), values(role, column))
            for other in names if other != column
        )]

    def set_choice(role, names, *, consequential=True):
        if len(names) > 1 and consequential:
            choices[role] = names
        mapping[role] = names[0] if names and (len(names) == 1 or not consequential) else None

    set_choice("date", candidates("date"))
    date = mapping["date"]
    if date and not frame.empty:
        valid_dates, _ = values("date", date)
        if any(value is not None for value in valid_dates):
            from input_precheck import summary_row_mask

            summary_rows = summary_row_mask(frame, mapping, date_format)
            active_rows = ~(summary_rows | frame.isna().all(axis=1))
        else:
            problems.append(f"「{date}」没有可解析的日期")
    for role in ("summary", "voucher_word", "voucher"):
        set_choice(role, candidates(role), consequential=False)
    set_choice("account", candidates("account"))

    amounts = {role: candidates(role) for role in ("debit", "credit", "amount")}
    balances = candidates("balance")
    directions = candidates("direction")
    split_mode = "debit_credit"
    single_modes = ("single_amount_with_direction", "signed_amount")
    if forced_mode is not None and forced_mode not in (split_mode, *single_modes):
        problems.append(f"未知金额模式「{forced_mode}」")
    if not forced_mode and "amount" in requested:
        forced_mode = single_modes[0] if directions else single_modes[1]
    if not forced_mode and {"debit", "credit"} & requested:
        forced_mode = split_mode

    def layouts():
        if forced_mode in (None, split_mode):
            for debit, credit in product(amounts["debit"], amounts["credit"]):
                tags = {amount_basis_label(debit), amount_basis_label(credit)} - {""}
                if len(tags) < 2 and "冲突" not in tags:
                    yield {"mode": split_mode, "debit": debit, "credit": credit}
        mode = forced_mode or ("single_amount_with_direction" if directions else "signed_amount")
        if mode in single_modes:
            selected_directions = directions if mode == "single_amount_with_direction" else [None]
            for amount, direction in product(amounts["amount"], selected_directions):
                yield {"mode": mode, "amount": amount, "direction": direction}

    def transactions(layout):
        mode = layout["mode"]
        if mode == "debit_credit":
            debit, d_present = values("debit", layout["debit"])
            credit, c_present = values("credit", layout["credit"])
            sign = 1 if source == "bank" else -1
            return tuple(
                ((c or 0) - (d or 0)) * sign
                if (dp or cp) and not ((dp and d is None) or (cp and c is None)) else None
                for d, c, dp, cp in zip(debit, credit, d_present, c_present)
            )
        amount, _ = values("amount", layout["amount"])
        if mode == "signed_amount":
            return amount
        directions_values, _ = values("direction", layout["direction"])
        return interpret_directions(amount, directions_values, balances)[0]

    def interpret_directions(amount, directions_values, balance_columns):
        signed = tuple(directed_amount(value, sign) for value, sign in zip(amount, directions_values))
        absolute = tuple(directed_amount(value, sign, absolute=True) for value, sign in zip(amount, directions_values))
        use_absolute = signed != absolute and any(
            _verified_chain(absolute, values("balance", balance)[0])
            and not _verified_chain(signed, values("balance", balance)[0])
            for balance in balance_columns
        )
        return (absolute if use_absolute else signed), use_absolute

    alternatives = list(layouts())
    if not alternatives:
        mode = forced_mode or ("single_amount_with_direction" if directions and amounts["amount"] else "signed_amount" if amounts["amount"] else "debit_credit")
        mapping["mode"] = mode
        required = ("debit", "credit") if mode == "debit_credit" else ("amount",)
        for role in required:
            set_choice(role, amounts[role])
            if not amounts[role]:
                problems.append(f"无法识别{FIELD_LABELS[role]}列")
        if mode == "debit_credit" and all(amounts[role] for role in required):
            problems.append("借方和贷方金额列的币种口径不一致")
        if mode == "single_amount_with_direction":
            set_choice("direction", directions)
            if not directions:
                problems.append("无法识别收支方向列")
        set_choice("balance", balances)
    else:
        if not frame.empty and len(alternatives) > 1:
            with_data = [layout for layout in alternatives if any(value is not None for value in transactions(layout))]
            alternatives = with_data or alternatives
        paired = [(layout, balance) for layout, balance in product(alternatives, balances or [None])
                  if not balance or _compatible_balance(layout, balance)]
        if not paired:
            paired = [(layout, None) for layout in alternatives]
            if "balance" in requested:
                problems.append("余额列和金额列的币种口径不一致")
        if len(paired) > 1 and active_rows is not None and any(balance for _, balance in paired):
            supported = [(layout, balance) for layout, balance in paired
                         if balance and _verified_chain(transactions(layout), values("balance", balance)[0])]
            if supported:
                paired = supported

        modes = {layout["mode"] for layout, _ in paired}
        if len(modes) > 1:
            # 同一套交易在原表既有分列又有净额时，保留可追溯的收付分列。
            equivalent = {(_layout_basis(layout), transactions(layout)) for layout, _ in paired}
            if len(equivalent) == 1:
                paired = [(layout, balance) for layout, balance in paired if layout["mode"] == "debit_credit"]
            else:
                choices["mode"] = sorted(modes)
        selected = [layout for layout, _ in paired]
        mapping["mode"] = selected[0]["mode"]
        for role in ("debit", "credit", "amount", "direction"):
            set_choice(role, sorted({layout[role] for layout in selected if layout.get(role)}))
        set_choice("balance", sorted({balance for _, balance in paired if balance}))
    for role in ("debit", "credit", "amount", "direction", "balance"):
        mapping.setdefault(role, None)
    if mapping["mode"] == "single_amount_with_direction" and mapping["amount"] and mapping["direction"]:
        mapping["direction_uses_absolute"] = interpret_directions(
            values("amount", mapping["amount"])[0],
            values("direction", mapping["direction"])[0],
            [mapping["balance"]] if mapping["balance"] else [],
        )[1]

    amount_keys = ("debit", "credit") if mapping["mode"] == "debit_credit" else ("amount",)
    tags = {amount_basis_label(mapping[key]) for key in amount_keys} - {""}
    basis = next(iter(tags)) if len(tags) == 1 and "冲突" not in tags else ""
    mapping["amount_basis"] = basis
    set_choice("currency", candidates("currency", basis))
    if mapping["currency"] and basis:
        tag = amount_basis_label(mapping["currency"])
        if tag and tag != basis:
            problems.append("币种列和金额列的口径不一致")
    if not date and "date" not in choices:
        problems.append("无法识别交易日期列")
    if overrides:
        problems.append("无法应用以下字段设置：" + "、".join(sorted(overrides)))
    return ColumnRecognition(mapping, choices, problems)


def _layout_basis(layout):
    return frozenset(amount_basis_label(layout.get(role)) for role in ("debit", "credit", "amount")) - {""}


def _compatible_balance(layout, balance):
    basis = _layout_basis(layout)
    balance_basis = amount_basis_label(balance)
    return not balance_basis or not basis or basis == {balance_basis}


def _more_complete(left, right):
    left_values, left_present = left
    right_values, right_present = right
    return (
        sum(value is not None for value in left_values) > sum(value is not None for value in right_values)
        and all(a == b or b is None for a, b in zip(left_values, right_values))
        and all(a is not None or not present or other is None and other_present
                for a, present, other, other_present in zip(left_values, left_present, right_values, right_present))
    )


def _verified_chain(amounts, balances):
    """用原表相邻交易验证差额，支持日期升序或降序的流水。"""
    for reverse in (False, True):
        rows = list(zip(amounts, balances))
        if reverse:
            rows.reverse()
        pairs = [(current - before, amount) for (_, before), (amount, current) in zip(rows, rows[1:])
                 if before is not None and current is not None]
        if sum(change != 0 for change, _ in pairs) >= 2 and all(change == amount for change, amount in pairs):
            return True
    return False
