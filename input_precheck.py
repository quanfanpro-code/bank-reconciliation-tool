"""银行流水和银行日记账的自包含输入预检查。"""

from __future__ import annotations

import re
import unicodedata
import json
from dataclasses import dataclass, field, replace
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from statistics import pstdev
from typing import Any, Iterable, Sequence

import pandas as pd

from data_structures import OverallControlResult
from utils import clean_excel_string


SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv"}
SUMMARY_ROW_KEYWORDS = (
    "本日合计", "本日累计", "本日发生额", "本日余额", "本日结存", "本日小计",
    "本旬合计", "本旬累计", "本旬发生额", "本旬余额", "本旬结存", "本旬小计",
    "本月合计", "本月累计", "本月发生额", "本月余额", "本月结存", "本月小计",
    "本季合计", "本季累计", "本季发生额", "本季余额", "本季结存", "本季小计",
    "本年合计", "本年累计", "本年发生额", "本年余额", "本年结存", "本年小计",
    "本期合计", "本期累计", "本期发生额", "本期余额", "本期结存", "本期小计",
    "日计", "月计", "季计", "年计", "期计", "日结", "月结", "季结", "年结",
    "合计", "累计", "总计", "小计", "大计", "发生额", "余额", "结存",
    "本页合计", "本页累计", "本页小计", "过次页", "承前页",
    "期初余额", "期末余额", "期初结存", "期末结存",
    "年初余额", "年末余额", "年初结存", "年末结存",
    "月初余额", "月末余额", "月初结存", "月末结存",
    "结转下年", "结转下期", "结转下月", "上年结转", "上期结转", "上月结转",
    "上年结余", "上期结余", "承前余额", "结转余额",
    "当前合计", "当前累计", "当前余额",
)
HEADER_KEYWORD_GROUPS = (
    ("日期", "date", "交易时间", "记账时间"),
    ("金额", "amount", "发生额"),
    ("借方", "借", "debit", "支出"),
    ("贷方", "贷", "credit", "收入"),
    ("方向", "借贷标志"),
    ("摘要", "summary", "用途", "业务说明"),
    ("余额", "balance"),
    ("对方户名", "对方名称", "附言", "备注"),
    ("凭证", "voucher"),
)


@dataclass(frozen=True)
class HeaderCandidate:
    """一个可追溯的表头范围候选。"""

    skiprows: int
    header_rows: int
    score: float
    columns: tuple[str, ...]
    evidence: tuple[str, ...] = ()


@dataclass(frozen=True)
class TableStructure:
    """实际采用的表头结构和备选证据。"""

    skiprows: int
    header_rows: int
    columns: list[str]
    score: float
    ambiguous: bool = False
    candidates: tuple[HeaderCandidate, ...] = field(default_factory=tuple)
    explanation: str = ""


@dataclass(frozen=True)
class PrecheckItem:
    """一项可展示、可留痕的输入检查结果。"""

    name: str
    bank_result: str
    journal_result: str
    comparison: str
    status: str
    explanation: str

    def as_dict(self) -> dict[str, str]:
        return {
            "检查项目": self.name,
            "银行流水结果": self.bank_result,
            "银行日记账结果": self.journal_result,
            "双方比较结果": self.comparison,
            "状态": self.status,
            "说明": self.explanation,
        }


@dataclass(frozen=True)
class InputPrecheckReport:
    """本次核对的全部输入检查结果。"""

    items: tuple[PrecheckItem, ...]
    source_info: tuple[dict[str, Any], ...] = ()
    population_rows: tuple[dict[str, Any], ...] = ()
    overall_control: OverallControlResult | None = None
    parse_errors: tuple[dict[str, Any], ...] = ()

    @property
    def has_blockers(self) -> bool:
        return any(item.status == "无法计算" for item in self.items)

    @property
    def has_warnings(self) -> bool:
        return any(item.status == "疑点" for item in self.items)

    def blocker_message(self) -> str:
        return "\n".join(
            f"• {item.name}：{item.explanation}"
            for item in self.items
            if item.status == "无法计算"
        )

    def warning_message(self) -> str:
        return "\n".join(
            f"• {item.name}：{item.explanation}"
            for item in self.items
            if item.status == "疑点"
        )

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame(
            [item.as_dict() for item in self.items],
            columns=(
                "检查项目",
                "银行流水结果",
                "银行日记账结果",
                "双方比较结果",
                "状态",
                "说明",
            ),
        )

    def source_dataframe(self) -> pd.DataFrame:
        columns = (
            "来源", "文件路径", "文件SHA256", "文件大小", "工作表",
            "表头起始行", "表头行数", "采用映射",
        )
        return pd.DataFrame(list(self.source_info), columns=columns)

    def population_dataframe(self) -> pd.DataFrame:
        columns = (
            "来源", "原文件行号", "唯一处置类别", "处置原因", "原日期列名",
            "原日期值", "原金额列名", "原金额值", "原方向列名", "原方向值",
            "原始可解析净额", "金额去向说明", "标准化净额", "进入匹配", "账户原值", "币种原值",
        )
        return pd.DataFrame(list(self.population_rows), columns=columns)


class InputPrecheckBlockedError(ValueError):
    """输入存在硬错误，正式匹配不得开始。"""

    def __init__(
        self,
        report: InputPrecheckReport,
        report_path: str | Path | None = None,
        report_write_error: str | None = None,
    ):
        self.report = report
        self.report_path = str(report_path) if report_path else None
        self.problem_report_path = self.report_path
        self.report_write_error = report_write_error
        super().__init__(report.blocker_message() or "输入预检查未通过")


def write_input_problem_report(
    report: InputPrecheckReport,
    output_path: str | Path,
) -> Path:
    """在正式匹配前写出独立的输入问题证据。"""
    from make_excel import make_excel

    path = Path(output_path)
    def safe_table(frame: pd.DataFrame) -> pd.DataFrame:
        result = frame.copy()
        for column in result.columns:
            result[column] = result[column].map(
                lambda value: clean_excel_string(value)
                if isinstance(value, str)
                else value
            )
        return result

    tables = [
            ("输入检查", safe_table(report.to_dataframe())),
            ("运行资料与映射", safe_table(report.source_dataframe())),
            ("数据入口处置", safe_table(report.population_dataframe())),
        ]
    if report.parse_errors:
        errors = pd.DataFrame(report.parse_errors).rename(columns={
            "type": "异常类型", "source_type": "来源", "row": "记录行号",
            "original_file_row": "原文件行号", "column": "字段", "original_value": "原值",
        })
        tables.append(("解析异常明细", safe_table(errors)))
    make_excel(
        tables,
        str(path),
        theme="deep-navy",
    )
    return path


def _cell_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"nan", "none"} else text


def _deduplicate(names: Sequence[str]) -> list[str]:
    counts: dict[str, int] = {}
    result = []
    for name in names:
        counts[name] = counts.get(name, 0) + 1
        result.append(name if counts[name] == 1 else f"{name}_{counts[name]}")
    return result


def flatten_header_rows(rows: Sequence[Sequence[Any]]) -> list[str]:
    """按列扁平化一至三行表头，保留原始文字顺序。"""
    if not rows:
        return []
    width = max(len(row) for row in rows)
    names = []
    for column_index in range(width):
        parts = []
        for row in rows:
            text = _cell_text(row[column_index] if column_index < len(row) else None)
            if text and (not parts or parts[-1] != text):
                parts.append(text)
        names.append("｜".join(parts) if parts else f"第{column_index + 1}列")
    return _deduplicate(names)


def _read_csv(path: Path, **kwargs: Any) -> pd.DataFrame:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "gbk", "latin1"):
        try:
            return pd.read_csv(path, encoding=encoding, **kwargs)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f"CSV 文件无法解码：{path.name}") from last_error


def _read_preview(path: Path, max_scan_rows: int) -> tuple[pd.DataFrame, list[tuple[int, int, int, int]]]:
    if path.suffix.lower() == ".csv":
        return (
            _read_csv(path, header=None, nrows=max_scan_rows, dtype=object),
            [],
        )

    preview = pd.read_excel(
        path,
        header=None,
        nrows=max_scan_rows,
        dtype=object,
    )
    merges: list[tuple[int, int, int, int]] = []
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=False, data_only=True)
        sheet = workbook.worksheets[0]
        merges = [
            (
                merged.min_row - 1,
                merged.max_row - 1,
                merged.min_col - 1,
                merged.max_col - 1,
            )
            for merged in sheet.merged_cells.ranges
            if merged.min_row <= max_scan_rows
        ]
        workbook.close()
    return preview, merges


def _expand_merges(
    preview: pd.DataFrame,
    merges: Iterable[tuple[int, int, int, int]],
) -> pd.DataFrame:
    expanded = preview.copy()
    for min_row, max_row, min_col, max_col in merges:
        if min_row >= len(expanded) or min_col >= len(expanded.columns):
            continue
        value = expanded.iat[min_row, min_col]
        for row in range(min_row, min(max_row + 1, len(expanded))):
            for column in range(min_col, min(max_col + 1, len(expanded.columns))):
                expanded.iat[row, column] = value
    return expanded


def _keyword_hits(values: Iterable[Any]) -> int:
    text = " ".join(_cell_text(value).lower() for value in values)
    return sum(
        1
        for group in HEADER_KEYWORD_GROUPS
        if any(keyword.lower() in text for keyword in group)
    )


def _looks_numeric_or_date(value: Any) -> bool:
    if pd.isna(value) or isinstance(value, (datetime, date, int, float)):
        return not pd.isna(value)
    text = str(value).strip().replace(",", "")
    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
        return True
    return bool(
        re.fullmatch(
            r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2}|\d{8}",
            text,
        )
    )


def _candidate_score(
    raw_preview: pd.DataFrame,
    expanded: pd.DataFrame,
    start: int,
    depth: int,
) -> HeaderCandidate | None:
    data_start = start + depth
    if data_start >= len(expanded):
        return None
    header = expanded.iloc[start:data_start]
    raw_header = raw_preview.iloc[start:data_start]
    columns = flatten_header_rows(header.values.tolist())
    if not columns:
        return None

    data = expanded.iloc[data_start:data_start + 5]
    data = data.loc[data.notna().any(axis=1)]
    if data.empty:
        return None

    width = len(expanded.columns)
    densities = [row.notna().sum() / width for _, row in data.iterrows()]
    coverage = sum(name and not name.startswith("第") for name in columns) / width
    keywords = _keyword_hits(header.to_numpy().ravel())
    transitions = 0
    for column_index in range(width):
        header_has_text = any(
            _cell_text(value) and not _looks_numeric_or_date(value)
            for value in header.iloc[:, column_index]
        )
        first_data_value = next(
            (
                value
                for value in data.iloc[:, column_index]
                if _cell_text(value)
            ),
            None,
        )
        if header_has_text and first_data_value is not None and _looks_numeric_or_date(first_data_value):
            transitions += 1

    raw_nonempty = raw_header.notna().sum(axis=1)
    title_penalty = 1 if int(raw_nonempty.iloc[0]) <= 2 else 0
    next_header_penalty = 1 if _keyword_hits(data.iloc[0].tolist()) >= 2 else 0
    data_like_header_ratio = sum(
        _looks_numeric_or_date(value)
        for value in header.iloc[-1]
        if _cell_text(value)
    ) / width
    stability = 1 - min(pstdev(densities), 1) if len(densities) > 1 else 1
    score = (
        keywords * 1.5
        + coverage * 2
        + (transitions / width) * 2
        + sum(densities) / len(densities)
        + stability
        - title_penalty * 4
        - next_header_penalty * 2
        - data_like_header_ratio * 5
        - (depth - 1) * 0.1
    )
    evidence = (
        f"业务词组{keywords}个",
        f"列覆盖率{coverage:.0%}",
        f"数据区稳定度{stability:.0%}",
        f"类型转换列{transitions}个",
        f"表头数据型占比{data_like_header_ratio:.0%}",
    )
    return HeaderCandidate(start, depth, round(score, 4), tuple(columns), evidence)


def detect_table_structure(
    file_path: str | Path,
    max_scan_rows: int = 40,
) -> TableStructure:
    """基于多类结构证据识别银行明细表头。"""
    path = Path(file_path)
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"文件不存在：{path}")
    if path.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ValueError(f"不支持的文件格式：{path.suffix or '无后缀'}")

    preview, merges = _read_preview(path, max_scan_rows)
    preview = preview.dropna(axis=1, how="all")
    if preview.empty or len(preview.columns) == 0:
        raise ValueError(f"文件没有可识别的表格内容：{path.name}")
    expanded = _expand_merges(preview, merges)

    candidates = []
    max_start = min(len(preview) - 1, max_scan_rows - 1)
    for start in range(max_start):
        for depth in range(1, min(3, len(preview) - start - 1) + 1):
            candidate = _candidate_score(preview, expanded, start, depth)
            if candidate is not None:
                candidates.append(candidate)
    if not candidates:
        raise ValueError(f"无法识别表头和数据区域：{path.name}")

    candidates.sort(key=lambda item: (-item.score, item.skiprows, item.header_rows))
    best = candidates[0]
    runner_up = next(
        (
            item
            for item in candidates[1:]
            if (item.skiprows, item.header_rows) != (best.skiprows, best.header_rows)
        ),
        None,
    )
    ambiguous = bool(runner_up and best.score - runner_up.score <= 0.35)
    explanation = (
        f"首选第{best.skiprows + 1}行起、共{best.header_rows}行表头"
        + ("；存在分数接近的备选结构" if ambiguous else "")
    )
    return TableStructure(
        skiprows=best.skiprows,
        header_rows=best.header_rows,
        columns=list(best.columns),
        score=best.score,
        ambiguous=ambiguous,
        candidates=tuple(candidates[:5]),
        explanation=explanation,
    )


def derive_header_columns(
    file_path: str | Path,
    skiprows: int,
    header_rows: int,
) -> list[str]:
    """按用户采用的表头范围派生稳定列名。"""
    if skiprows < 0 or header_rows < 1:
        raise ValueError("跳过行数不得小于0，表头行数必须大于等于1")
    path = Path(file_path)
    preview, merges = _read_preview(path, skiprows + header_rows)
    preview = preview.dropna(axis=1, how="all")
    if len(preview) < skiprows + header_rows:
        raise ValueError("指定的表头范围超出文件内容")
    expanded = _expand_merges(preview, merges)
    return flatten_header_rows(
        expanded.iloc[skiprows:skiprows + header_rows].values.tolist()
    )


def _required_columns(mapping: dict[str, Any]) -> list[tuple[str, str]]:
    required = [("date", "日期")]
    mode = mapping.get("mode", "debit_credit")
    if mode == "debit_credit":
        required.extend((("debit", "借方/支出"), ("credit", "贷方/收入")))
    elif mode == "single_amount_with_direction":
        required.extend((("amount", "金额"), ("direction", "方向")))
    elif mode == "signed_amount":
        required.append(("amount", "金额"))
    else:
        required.append(("mode", "金额模式"))
    return required


def _mapping_problems(mapping: dict[str, Any], columns: Sequence[Any]) -> list[str]:
    available = set(columns)
    selected = []
    problems = []
    for key, label in _required_columns(mapping):
        value = mapping.get(key)
        if key == "mode" or value in (None, "", "(无)"):
            problems.append(f"缺少{label}列")
        elif value not in available:
            problems.append(f"{label}列不存在：{value}")
        elif value in selected:
            problems.append(f"{label}列与其他必填映射重复：{value}")
        else:
            selected.append(value)
    for key, label in (
        ("account", "本方账号"),
        ("currency", "币种"),
        ("voucher_word", "凭证字/类型"),
        ("voucher", "凭证号"),
        ("balance", "余额"),
    ):
        value = mapping.get(key)
        if value not in (None, "", "(无)") and value not in available:
            problems.append(f"{label}列不存在：{value}")
    return problems


def _structure_result(structure: TableStructure) -> str:
    return (
        f"第{structure.skiprows + 1}行起，{structure.header_rows}行表头；"
        f"派生列名：{'、'.join(structure.columns)}"
    )


def _ambiguity_changes_mapping(
    structure: TableStructure,
    mapping: dict[str, Any],
) -> bool:
    if not structure.ambiguous:
        return False
    required_names = {
        mapping.get(key)
        for key, _label in _required_columns(mapping)
        if key != "mode" and mapping.get(key)
    }
    alternatives = [
        candidate
        for candidate in structure.candidates
        if (candidate.skiprows, candidate.header_rows)
        != (structure.skiprows, structure.header_rows)
    ]
    return not alternatives or any(
        not required_names.issubset(set(candidate.columns))
        for candidate in alternatives[:2]
    )


def _valid_dates(frame: pd.DataFrame) -> pd.Series:
    if frame.empty or "date" not in frame.columns:
        return pd.Series(dtype="datetime64[ns]")
    return pd.to_datetime(frame["date"], errors="coerce").dropna()


def _amount_totals(frame: pd.DataFrame) -> tuple[Decimal, Decimal] | None:
    if frame.empty or "amount" not in frame.columns:
        return None
    amounts = [
        Decimal(str(value))
        for value in frame["amount"]
        if pd.notna(value)
    ]
    if not amounts:
        return None
    income = sum((value for value in amounts if value > 0), Decimal("0"))
    expense = sum((-value for value in amounts if value < 0), Decimal("0"))
    return income, expense


def _error_count(
    parse_errors: Sequence[dict[str, Any]],
    source: str,
    error_type: str,
) -> int:
    return sum(
        1
        for error in parse_errors
        if error.get("source_type") == source and error.get("type") == error_type
    )


NON_TRANSACTION_KEYWORDS = (
    "合计",
    "累计",
    "小计",
    "总计",
    "日计",
    "月计",
    "年计",
    "期初",
    "期末",
    "承前页",
    "过次页",
    "统计",
)
NOTE_PREFIXES = ("注：", "说明：", "备注：", "单位：", "制表：")


def _non_transaction_mask(
    frame: pd.DataFrame,
    mapping: dict[str, Any],
) -> pd.Series:
    if frame.empty:
        return pd.Series(False, index=frame.index)
    empty = frame.map(lambda value: not _cell_text(value)).all(axis=1)
    combined = frame.apply(
        lambda row: " ".join(_cell_text(value) for value in row),
        axis=1,
    )
    generic_summary_or_note = combined.str.contains(
        "|".join(re.escape(word) for word in NON_TRANSACTION_KEYWORDS),
        na=False,
    ) | combined.str.startswith(NOTE_PREFIXES)
    column_names = {str(column).strip() for column in frame.columns}
    repeated_header = frame.apply(
        lambda row: sum(
            _cell_text(value) in column_names
            for value in row
            if _cell_text(value)
        )
        >= min(2, len(column_names)),
        axis=1,
    )
    date_column = mapping.get("date")
    dates = (
        pd.to_datetime(frame[date_column], errors="coerce")
        if date_column in frame.columns
        else pd.Series(pd.NaT, index=frame.index)
    )
    summary_column = mapping.get("summary")
    explicit_summary = (
        frame[summary_column].map(_cell_text).isin(SUMMARY_ROW_KEYWORDS)
        if summary_column in frame.columns
        else pd.Series(False, index=frame.index)
    )
    summary = explicit_summary | (generic_summary_or_note & dates.isna())
    nonempty_counts = frame.map(lambda value: bool(_cell_text(value))).sum(axis=1)
    title_or_note = (nonempty_counts == 1) & dates.isna()
    return empty | summary | repeated_header | title_or_note


def _auxiliary_result(
    frame: pd.DataFrame,
    mapping: dict[str, Any],
) -> tuple[str, bool]:
    selected = []
    for column in [mapping.get("summary"), *(mapping.get("auxiliary_text_columns") or [])]:
        if column and column not in selected:
            selected.append(column)
    usable = [column for column in selected if column in frame.columns]
    if not usable:
        return "未选择可用辅助文字列", True
    transaction_rows = frame.loc[~_non_transaction_mask(frame, mapping)]
    denominator = len(transaction_rows)
    if denominator == 0:
        return "没有可计算非空率的交易行", True
    parts = []
    low = False
    for column in usable:
        nonempty = transaction_rows[column].map(lambda value: bool(_cell_text(value))).sum()
        rate = nonempty / denominator
        parts.append(f"{column} {rate:.1%}")
        low = low or rate < 0.8
    missing = [column for column in selected if column not in frame.columns]
    if missing:
        parts.append(f"未找到：{'、'.join(missing)}")
        low = True
    return "；".join(parts), low


def _scope_values(
    frame: pd.DataFrame,
    mapping: dict[str, Any],
    identity: str,
) -> tuple[str, set[str], int, int]:
    """提取核对账户或币种，同时统计有效交易中的空白身份行。"""
    columns = []
    explicit_column = mapping.get(identity)
    if explicit_column and explicit_column in frame.columns:
        columns = [explicit_column]
    else:
        for candidate in frame.columns:
            name = _cell_text(candidate)
            lowered = name.lower()
            if identity == "account":
                matched = (
                    "对方" not in name
                    and (
                        "本方账号" in name
                        or name in {"账号", "银行账号", "账户账号", "卡号"}
                        or lowered in {"account", "account number", "account_number"}
                    )
                )
            else:
                matched = (
                    "对方" not in name
                    and (
                        "币种" in name
                        or "币别" in name
                        or lowered in {"currency", "currency code", "currency_code"}
                    )
                )
            if matched:
                columns.append(candidate)
    transaction_rows = frame.loc[~_non_transaction_mask(frame, mapping)]
    total_rows = len(transaction_rows)
    if not columns:
        return "未提供", set(), total_rows, total_rows

    def row_raw_values(row: pd.Series) -> set[str]:
        return {
            _cell_text(row[column]).upper()
            for column in columns
            if _cell_text(row[column])
        }

    row_values = transaction_rows.apply(row_raw_values, axis=1)
    raw_values = set().union(*row_values) if len(row_values) else set()
    missing_rows = int(row_values.map(lambda values: not values).sum())
    if not raw_values:
        return f"{'、'.join(map(str, columns))}未提供有效值", set(), missing_rows, total_rows

    def normalize(value: str) -> str:
        normalized = unicodedata.normalize("NFKC", value).strip().upper()
        if identity == "currency":
            aliases = {
                "人民币": "CNY",
                "人民币元": "CNY",
                "RMB": "CNY",
                "CNY": "CNY",
                "156": "CNY",
            }
            return aliases.get(normalized, normalized)
        compact = re.sub(r"[\s\-_]", "", normalized)
        # Excel 常把纯数字账号读成 12345.0；仅去掉确定无意义的小数尾零。
        if re.fullmatch(r"\d+\.0+", compact):
            return compact.split(".", 1)[0]
        return compact

    values = {normalize(value) for value in raw_values}
    result = f"{'、'.join(map(str, columns))}：{'、'.join(sorted(raw_values))}"
    if missing_rows:
        result += f"；{total_rows}行有效交易中缺失{missing_rows}行"
    return result, values, missing_rows, total_rows


def _scope_item(
    *,
    name: str,
    identity: str,
    raw_bank: pd.DataFrame,
    raw_journal: pd.DataFrame,
    bank_mapping: dict[str, Any],
    journal_mapping: dict[str, Any],
) -> PrecheckItem:
    missing_columns = [
        f"{label}显式选择的列“{mapping[identity]}”不存在"
        for label, frame, mapping in (
            ("银行流水", raw_bank, bank_mapping), ("银行日记账", raw_journal, journal_mapping)
        )
        if mapping.get(identity) and mapping[identity] not in frame.columns
    ]
    if missing_columns:
        return PrecheckItem(name, "显式映射待修正", "显式映射待修正", "范围映射无效", "无法计算",
                            "；".join(missing_columns) + "；程序已停止匹配，不能改用其他列。")
    bank_result, bank_values, bank_missing, bank_total = _scope_values(
        raw_bank, bank_mapping, identity
    )
    journal_result, journal_values, journal_missing, journal_total = _scope_values(
        raw_journal, journal_mapping, identity
    )
    missing = (
        not bank_values
        or not journal_values
        or bank_missing > 0
        or journal_missing > 0
    )
    mixed = len(bank_values) > 1 or len(journal_values) > 1
    mismatch = bool(bank_values and journal_values and bank_values != journal_values)
    if mixed:
        comparison = "单份文件存在多个取值"
        explanation = f"{name}范围混合，无法确认本次核对范围；程序已停止匹配。"
        status = "无法计算"
    elif mismatch:
        comparison = "双方不一致"
        explanation = f"双方{name}不一致，无法确认核对范围；程序已停止匹配。"
        status = "无法计算"
    elif missing:
        comparison = "范围身份未完全验证"
        explanation = (
            f"{name}字段缺失或无有效值：银行流水缺失{bank_missing}/{bank_total}行，"
            f"银行日记账缺失{journal_missing}/{journal_total}行；程序继续处理，并在报告中保留范围限制。"
        )
        status = "疑点"
    else:
        comparison = "双方一致"
        explanation = f"双方{name}一致。"
        status = "通过"
    return PrecheckItem(
        name=name,
        bank_result=bank_result,
        journal_result=journal_result,
        comparison=comparison,
        status=status,
        explanation=explanation,
    )


def _amount_basis(mapping: dict[str, Any]) -> tuple[str, str, bool]:
    """返回金额口径说明、规范口径及单侧映射是否自相矛盾。"""
    aliases = {
        "原币": "原币",
        "外币": "原币",
        "交易币": "原币",
        "ORIGINAL": "原币",
        "ORIGINAL_CURRENCY": "原币",
        "TRANSACTION_CURRENCY": "原币",
        "本位币": "本位币",
        "本币": "本位币",
        "记账本位币": "本位币",
        "BASE": "本位币",
        "BASE_CURRENCY": "本位币",
        "FUNCTIONAL_CURRENCY": "本位币",
    }
    mode = mapping.get("mode", "debit_credit")
    keys = ("debit", "credit") if mode == "debit_credit" else ("amount",)
    columns = [str(mapping.get(key)) for key in keys if mapping.get(key)]
    detected = set()
    for column in columns:
        name = unicodedata.normalize("NFKC", column)
        if re.search(r"原币|外币|交易币", name):
            detected.add("原币")
        if re.search(r"本位币|记账本位币|本币", name):
            detected.add("本位币")
    column_text = "、".join(columns) if columns else "未选择金额列"
    if len(detected) > 1:
        return f"映射列混用原币和本位币：{column_text}", "", True

    explicit = _cell_text(mapping.get("amount_basis"))
    if explicit:
        normalized = unicodedata.normalize("NFKC", explicit).strip().upper()
        value = aliases.get(normalized, "")
        if not value:
            return f"未识别的显式口径：{explicit}", "", True
        if detected and value != next(iter(detected)):
            detected_value = next(iter(detected))
            return (
                f"显式口径{value}与金额列标记{detected_value}矛盾：{column_text}",
                "",
                True,
            )
        return f"{value}（映射明确指定）", value, False

    if detected:
        value = next(iter(detected))
        return f"{value}（列：{column_text}）", value, False
    return f"未注明（列：{column_text}）", "", False


def _amount_basis_item(
    bank_mapping: dict[str, Any],
    journal_mapping: dict[str, Any],
) -> PrecheckItem:
    bank_result, bank_basis, bank_conflict = _amount_basis(bank_mapping)
    journal_result, journal_basis, journal_conflict = _amount_basis(journal_mapping)
    if bank_conflict or journal_conflict:
        comparison = "至少一侧金额列口径自相矛盾"
        status = "无法计算"
        explanation = "当前映射混用或错误声明原币、本位币，程序已停止匹配。"
    elif bank_basis and journal_basis and bank_basis != journal_basis:
        comparison = f"银行流水{bank_basis} / 银行日记账{journal_basis}"
        status = "无法计算"
        explanation = "双方采用的金额口径不一致，程序已停止匹配。"
    elif bank_basis and journal_basis:
        comparison = f"双方均为{bank_basis}"
        status = "通过"
        explanation = f"双方金额列均按{bank_basis}口径核对。"
    else:
        comparison = "至少一侧金额列未注明原币或本位币"
        status = "疑点"
        explanation = "无法仅凭通用列名验证原币或本位币口径；程序继续处理并在报告中保留范围限制。"
    return PrecheckItem(
        name="金额口径",
        bank_result=bank_result,
        journal_result=journal_result,
        comparison=comparison,
        status=status,
        explanation=explanation,
    )


def _overall_balance_item(control: OverallControlResult) -> PrecheckItem:
    """把余额总体控制转成输入检查表的一项，不把未提供余额伪装成已核对。"""
    def amount(value: Decimal | None) -> str:
        return "未取得" if value is None else f"{value:.2f}"

    def side_result(source: str) -> str:
        if source == "bank":
            status = control.bank_balance_status
            initial = control.bank_initial_balance
            ending = control.bank_ending_balance
            expected = control.bank_expected_ending_balance
            difference = control.bank_balance_diff
            label = "银行流水"
        else:
            status = control.journal_balance_status
            initial = control.journal_initial_balance
            ending = control.journal_ending_balance
            expected = control.journal_expected_ending_balance
            difference = control.journal_balance_diff
            label = "银行日记账"
        if status == "未实施":
            return "未提供可用余额，余额核对未实施"
        anomalies = sum(
            str(item.get("来源", "")) == label
            for item in control.continuity_anomalies
        )
        return (
            f"期初{amount(initial)}；预期期末{amount(expected)}；"
            f"实际期末{amount(ending)}；控制差额{amount(difference)}；"
            f"连续性异常{anomalies}处；"
            + "；".join(reason for reason in control.reasons if reason.startswith(label))
        )

    balance_anomaly = (
        control.bank_balance_status == "疑点"
        or control.journal_balance_status == "疑点"
        or bool(control.continuity_anomalies)
        or (control.initial_balance_diff is not None and control.initial_balance_diff != 0)
        or (control.ending_balance_diff is not None and control.ending_balance_diff != 0)
    )
    missing_side = (
        control.bank_balance_status == "未实施"
        or control.journal_balance_status == "未实施"
    )
    if control.balance_check_possible:
        comparison = (
            f"期初余额差额{amount(control.initial_balance_diff)}；"
            f"期末余额差额{amount(control.ending_balance_diff)}"
        )
    else:
        comparison = "至少一侧无可用余额，双方余额比较未实施"
    status = "疑点" if balance_anomaly or missing_side or control.scope_limited else "通过"
    explanation = (
        "总体控制范围受限："
        + "；".join(control.reasons)
        + "；相关候选不得无保留自动确认。"
        if control.scope_limited
        else (
            "余额连续性或期初加期间净发生额等于期末余额的控制存在异常，"
            "相关候选不得无保留自动确认。"
            if balance_anomaly
            else (
                "至少一侧未提供逐笔余额，程序未据此降低单笔关系，但总体余额核对范围受限。"
                if missing_side
                else "双方余额连续性及期初、期间发生额、期末余额控制通过。"
            )
        )
    )
    return PrecheckItem(
        name="总体余额控制",
        bank_result=side_result("bank"),
        journal_result=side_result("journal"),
        comparison=comparison,
        status=status,
        explanation=explanation,
    )


def _raw_population_amount(row: pd.Series, mapping: dict[str, Any], source: str) -> Decimal | None:
    """按标准化同一口径重算原行金额，无法解释的原值保留为空，不冒充零。"""
    from data_loader import parse_source_amount, direction_sign

    mode = mapping.get("mode", "debit_credit")
    if mode == "debit_credit":
        values = [row.get(mapping.get(key)) for key in ("debit", "credit")]
        if not any(_cell_text(value) for value in values):
            return None
        amounts = [parse_source_amount(value, source, False) if _cell_text(value) else Decimal("0")
                   for value in values]
        if any(value is None for value in amounts):
            return None
        debit, credit = amounts
        return credit - debit if source == "bank" else debit - credit
    amount = parse_source_amount(row.get(mapping.get("amount")), source, mode == "signed_amount")
    if amount is None:
        return None
    if mode == "single_amount_with_direction":
        sign = direction_sign(row.get(mapping.get("direction")), source)
        return amount * sign if sign is not None else None
    return amount


def _population_detail_rows(
    *,
    source_type: str,
    raw: pd.DataFrame,
    standardized: pd.DataFrame,
    mapping: dict[str, Any],
    structure: TableStructure,
    parse_errors: Sequence[dict[str, Any]],
) -> list[dict[str, Any]]:
    """将每个原始数据行唯一归入有效、非交易、解析异常或其他排除。"""
    valid_rows = {
        int(value)
        for value in standardized.get("original_file_row", pd.Series(dtype=object))
        if pd.notna(value)
    }
    issues_by_row: dict[int, set[str]] = {}
    for error in parse_errors:
        error_source = error.get("source_type")
        if error_source is None and error.get("type") == "被丢弃的汇总行":
            error_source = "journal"
        if error_source != source_type:
            continue
        row = error.get("original_file_row", error.get("row"))
        try:
            row_number = int(row)
        except (TypeError, ValueError):
            continue
        issues_by_row.setdefault(row_number, set()).add(
            str(error.get("type") or error.get("error_type") or "解析异常")
        )

    non_transaction = _non_transaction_mask(raw, mapping)
    date_column = mapping.get("date")
    amount_keys = (
        ("debit", "credit")
        if mapping.get("mode", "debit_credit") == "debit_credit"
        else ("amount",)
    )
    amount_columns = [mapping.get(key) for key in amount_keys if mapping.get(key)]
    direction_column = mapping.get("direction")

    standardized_amounts = {}
    if "original_file_row" in standardized.columns:
        for _, std_row in standardized.iterrows():
            try:
                row_number = int(std_row["original_file_row"])
            except (TypeError, ValueError):
                continue
            standardized_amounts[row_number] = std_row.get("amount")

    def matching_columns(identity: str) -> list[Any]:
        explicit = mapping.get(identity)
        if explicit:
            return [explicit] if explicit in raw.columns else []
        result = []
        for column in raw.columns:
            name = _cell_text(column)
            lowered = name.lower()
            if identity == "account":
                matched = "对方" not in name and (
                    "本方账号" in name
                    or name in {"账号", "银行账号", "账户账号", "卡号"}
                    or lowered in {"account", "account number", "account_number"}
                )
            else:
                matched = "对方" not in name and (
                    "币种" in name
                    or "币别" in name
                    or lowered in {"currency", "currency code", "currency_code"}
                )
            if matched:
                result.append(column)
        return result

    account_columns = matching_columns("account")
    currency_columns = matching_columns("currency")
    rows = []
    for position, (index, raw_row) in enumerate(raw.iterrows()):
        if "__file_row__" in raw.columns and pd.notna(raw_row.get("__file_row__")):
            file_row = int(raw_row["__file_row__"])
        else:
            file_row = position + int(structure.skiprows) + int(structure.header_rows) + 1
        if file_row in valid_rows:
            category = "有效交易"
            reasons = "日期和金额可用，已进入匹配"
        elif bool(non_transaction.loc[index]):
            category = "非交易行"
            reasons = "合计、累计、标题、重复表头、注释或空行"
        elif file_row in issues_by_row:
            category = "解析异常"
            reasons = "、".join(sorted(issues_by_row[file_row]))
        else:
            category = "其他排除"
            reasons = "未进入标准化交易且未被现有规则解释"

        amount_values = {
            str(column): raw_row.get(column)
            for column in amount_columns
            if column in raw.columns and _cell_text(raw_row.get(column))
        }
        account_values = [
            _cell_text(raw_row.get(column)) for column in account_columns
            if _cell_text(raw_row.get(column))
        ]
        currency_values = [
            _cell_text(raw_row.get(column)) for column in currency_columns
            if _cell_text(raw_row.get(column))
        ]
        original_net = _raw_population_amount(raw_row, mapping, source_type)
        rows.append({
            "来源": "银行流水" if source_type == "bank" else "银行日记账",
            "原文件行号": file_row,
            "唯一处置类别": category,
            "处置原因": reasons,
            "原日期列名": date_column or "",
            "原日期值": _cell_text(raw_row.get(date_column)) if date_column in raw.columns else "",
            "原金额列名": "、".join(map(str, amount_columns)),
            "原金额值": json.dumps(amount_values, ensure_ascii=False, default=str),
            "原方向列名": direction_column or "",
            "原方向值": _cell_text(raw_row.get(direction_column)) if direction_column in raw.columns else "",
            "原始可解析净额": original_net,
            "金额去向说明": (f"原行金额归入{category}，未与其他类别重复累计" if original_net is not None
                           else "原金额或方向无法解析或未提供；保留原值，不以零计入金额勾稽"),
            "标准化净额": standardized_amounts.get(file_row),
            "进入匹配": category == "有效交易",
            "账户原值": "、".join(account_values),
            "币种原值": "、".join(currency_values),
        })
    return rows


def build_input_precheck(
    *,
    raw_bank: pd.DataFrame,
    raw_journal: pd.DataFrame,
    bank: pd.DataFrame,
    journal: pd.DataFrame,
    bank_mapping: dict[str, Any],
    journal_mapping: dict[str, Any],
    bank_structure: TableStructure,
    journal_structure: TableStructure,
    parse_errors: Sequence[dict[str, Any]] = (),
    source_info: Sequence[dict[str, Any]] = (),
    overall_control: OverallControlResult | None = None,
) -> InputPrecheckReport:
    """对已读取和标准化的双方数据执行统一检查。"""
    if overall_control is None:
        # 局部导入，避免 DataLoader -> input_precheck -> balance -> DataLoader 的循环。
        from balance import build_overall_controls
        overall_control = build_overall_controls(bank, journal)
    items = []

    file_blocked = raw_bank.empty or raw_journal.empty
    items.append(
        PrecheckItem(
            "文件读取",
            f"可读取，{len(raw_bank)}行" if not raw_bank.empty else "没有可核对数据",
            f"可读取，{len(raw_journal)}行" if not raw_journal.empty else "没有可核对数据",
            "双方均已读取" if not file_blocked else "至少一侧没有数据",
            "无法计算" if file_blocked else "通过",
            "文件没有可核对数据，请检查表头和数据区域。" if file_blocked else "文件可正常读取。",
        )
    )

    bank_major = _ambiguity_changes_mapping(bank_structure, bank_mapping)
    journal_major = _ambiguity_changes_mapping(journal_structure, journal_mapping)
    ambiguous = bank_structure.ambiguous or journal_structure.ambiguous
    user_override = any(
        structure.explanation.startswith("采用用户设置")
        for structure in (bank_structure, journal_structure)
    )
    structure_status = (
        "疑点"
        if bank_major or journal_major
        else ("疑点" if ambiguous or user_override else "通过")
    )
    items.append(
        PrecheckItem(
            "表格结构",
            _structure_result(bank_structure),
            _structure_result(journal_structure),
            (
                "存在重大歧义"
                if bank_major or journal_major
                else (
                    "采用用户设置"
                    if user_override
                    else ("存在备选结构" if ambiguous else "结构明确")
                )
            ),
            structure_status,
            (
                "表头候选会改变必填列映射，请返回确认表头位置和层级。"
                if bank_major or journal_major
                else (
                    "当前采用用户设置的表头范围，与程序首选候选不同；列映射仍然有效。"
                    if user_override
                    else ("检测到分数接近的表头候选，请确认当前列映射。" if ambiguous else "表头和数据区域可用。")
                )
            ),
        )
    )

    bank_dates = _valid_dates(bank)
    journal_dates = _valid_dates(journal)
    date_blocked = bank_dates.empty or journal_dates.empty
    bank_date_errors = sum(
        _error_count(parse_errors, "bank", error_type)
        for error_type in ("日期解析失败", "空日期行")
    )
    journal_date_errors = sum(
        _error_count(parse_errors, "journal", error_type)
        for error_type in ("日期解析失败", "空日期行")
    )
    if date_blocked:
        date_status = "无法计算"
        date_explanation = "至少一侧日期全部无法解析，请检查日期列和日期格式。"
        date_comparison = "无法比较"
    else:
        bank_range = (bank_dates.min(), bank_dates.max())
        journal_range = (journal_dates.min(), journal_dates.max())
        mismatch = bank_range != journal_range
        date_status = "疑点" if mismatch or bank_date_errors or journal_date_errors else "通过"
        date_comparison = "范围一致" if not mismatch else "范围不完全一致"
        date_explanation = (
            f"少量日期或金额解析失败：银行流水日期{bank_date_errors}行，"
            f"银行日记账日期{journal_date_errors}行。"
            if bank_date_errors or journal_date_errors
            else ("双方日期范围不完全一致，请确认是否属于正常未达期间。" if mismatch else "双方日期范围一致。")
        )
    items.append(
        PrecheckItem(
            "日期范围",
            "无法形成日期范围" if bank_dates.empty else f"{bank_dates.min():%Y-%m-%d} 至 {bank_dates.max():%Y-%m-%d}",
            "无法形成日期范围" if journal_dates.empty else f"{journal_dates.min():%Y-%m-%d} 至 {journal_dates.max():%Y-%m-%d}",
            date_comparison,
            date_status,
            date_explanation,
        )
    )

    items.append(
        _scope_item(
            name="核对账户",
            identity="account",
            raw_bank=raw_bank,
            raw_journal=raw_journal,
            bank_mapping=bank_mapping,
            journal_mapping=journal_mapping,
        )
    )
    items.append(
        _scope_item(
            name="核对币种",
            identity="currency",
            raw_bank=raw_bank,
            raw_journal=raw_journal,
            bank_mapping=bank_mapping,
            journal_mapping=journal_mapping,
        )
    )

    items.append(_amount_basis_item(bank_mapping, journal_mapping))

    scope_warnings = [
        item
        for item in items
        if item.name in {"核对账户", "核对币种", "金额口径"}
        and item.status == "疑点"
    ]
    if scope_warnings:
        scope_reasons = tuple(
            f"{item.name}未完全验证：{item.comparison}"
            for item in scope_warnings
        )
        overall_control = replace(
            overall_control,
            scope_limited=True,
            reasons=tuple(
                dict.fromkeys((*overall_control.reasons, *scope_reasons))
            ),
        )

    bank_direction_errors = _error_count(parse_errors, "bank", "方向解析失败")
    journal_direction_errors = _error_count(parse_errors, "journal", "方向解析失败")
    direction_blocked = bool(bank_direction_errors or journal_direction_errors)
    items.append(
        PrecheckItem(
            "金额方向",
            f"银行口径；无法识别{bank_direction_errors}行",
            f"日记账口径；无法识别{journal_direction_errors}行",
            "贷增借减 / 借增贷减",
            "疑点" if direction_blocked else "通过",
            (
                "方向列存在无法识别的值，收入和支出方向不可靠。"
                if direction_blocked
                else "银行流水按贷增借减，银行日记账按借增贷减。"
            ),
        )
    )

    bank_totals = _amount_totals(bank)
    journal_totals = _amount_totals(journal)
    amount_blocked = bank_totals is None or journal_totals is None
    bank_amount_errors = _error_count(parse_errors, "bank", "金额解析失败")
    journal_amount_errors = _error_count(parse_errors, "journal", "金额解析失败")
    if amount_blocked:
        amount_status = "无法计算"
        amount_comparison = "无法比较"
        amount_explanation = "至少一侧金额全部无法解析，请检查金额列或借贷列。"
    else:
        income_diff = bank_totals[0] - journal_totals[0]
        expense_diff = bank_totals[1] - journal_totals[1]
        has_diff = income_diff != 0 or expense_diff != 0
        amount_status = "疑点" if has_diff or bank_amount_errors or journal_amount_errors else "通过"
        amount_comparison = f"收入差额 {income_diff:.2f}；支出差额 {expense_diff:.2f}"
        amount_explanation = (
            f"少量日期或金额解析失败：银行流水金额{bank_amount_errors}行，"
            f"银行日记账金额{journal_amount_errors}行。"
            if bank_amount_errors or journal_amount_errors
            else ("双方收入或支出合计存在差额；差额是核对对象，不直接视为输入错误。" if has_diff else "双方收入和支出合计一致。")
        )
    items.append(
        PrecheckItem(
            "金额合计",
            "无法形成金额合计" if bank_totals is None else f"收入 {bank_totals[0]:.2f}；支出 {bank_totals[1]:.2f}",
            "无法形成金额合计" if journal_totals is None else f"收入 {journal_totals[0]:.2f}；支出 {journal_totals[1]:.2f}",
            amount_comparison,
            amount_status,
            amount_explanation,
        )
    )
    bank_balance_errors = _error_count(parse_errors, "bank", "余额解析失败")
    journal_balance_errors = _error_count(parse_errors, "journal", "余额解析失败")
    if bank_balance_errors or journal_balance_errors:
        balance_parse_reason = (
            "余额解析失败："
            f"银行流水{bank_balance_errors}行，银行日记账{journal_balance_errors}行；"
            "总体余额控制只覆盖其余可解析行。"
        )
        overall_control = replace(
            overall_control,
            scope_limited=True,
            reasons=tuple(
                dict.fromkeys((*overall_control.reasons, balance_parse_reason))
            ),
        )
    items.append(_overall_balance_item(overall_control))

    bank_non_transactions = int(_non_transaction_mask(raw_bank, bank_mapping).sum())
    journal_non_transactions = int(_non_transaction_mask(raw_journal, journal_mapping).sum())
    has_non_transactions = bank_non_transactions > 0 or journal_non_transactions > 0
    items.append(
        PrecheckItem(
            "非交易行",
            f"识别并排除{bank_non_transactions}行",
            f"识别并排除{journal_non_transactions}行",
            f"合计{bank_non_transactions + journal_non_transactions}行",
            "疑点" if has_non_transactions else "通过",
            "检测到合计、累计、统计、标题、重复表头、注释或空行。" if has_non_transactions else "未发现混入数据区的非交易行。",
        )
    )
    population_rows = _population_detail_rows(
        source_type="bank", raw=raw_bank, standardized=bank,
        mapping=bank_mapping, structure=bank_structure, parse_errors=parse_errors,
    ) + _population_detail_rows(
        source_type="journal", raw=raw_journal, standardized=journal,
        mapping=journal_mapping, structure=journal_structure, parse_errors=parse_errors,
    )
    population_counts = {
        source: {
            category: sum(
                row["来源"] == source and row["唯一处置类别"] == category
                for row in population_rows
            )
            for category in ("有效交易", "非交易行", "解析异常", "其他排除")
        }
        for source in ("银行流水", "银行日记账")
    }
    bank_parse_errors = population_counts["银行流水"]["解析异常"]
    journal_parse_errors = population_counts["银行日记账"]["解析异常"]
    bank_unexplained = population_counts["银行流水"]["其他排除"]
    journal_unexplained = population_counts["银行日记账"]["其他排除"]
    population_risk = bool(
        bank_parse_errors
        or journal_parse_errors
        or bank_unexplained
        or journal_unexplained
    )
    if population_risk:
        population_control_reasons = []
        if bank_parse_errors or journal_parse_errors:
            population_control_reasons.append(
                "数据入口存在解析异常："
                f"银行流水{bank_parse_errors}行，银行日记账{journal_parse_errors}行"
            )
        if bank_unexplained or journal_unexplained:
            population_control_reasons.append(
                "数据入口存在其他排除："
                f"银行流水{bank_unexplained}行，银行日记账{journal_unexplained}行"
            )
        combined_reasons = tuple(
            dict.fromkeys((*overall_control.reasons, *population_control_reasons))
        )
        overall_control = replace(
            overall_control,
            scope_limited=True,
            reasons=combined_reasons,
        )
        items = [
            _overall_balance_item(overall_control)
            if item.name == "总体余额控制"
            else item
            for item in items
        ]

    def population_summary(source: str) -> str:
        rows = [row for row in population_rows if row["来源"] == source]
        counts = population_counts[source]
        raw_net = sum((row["原始可解析净额"] for row in rows if row["原始可解析净额"] is not None), Decimal("0"))
        category_net = {
            category: sum((row["标准化净额"] if category == "有效交易" else row["原始可解析净额"]
                           for row in rows if row["唯一处置类别"] == category
                           and (row["标准化净额"] if category == "有效交易" else row["原始可解析净额"]) is not None), Decimal("0"))
            for category in counts
        }
        unknown = sum(row["原始可解析净额"] is None for row in rows)
        diff = raw_net - sum(category_net.values(), Decimal("0"))
        return (
            f"原始{len(rows)}行；有效交易{counts['有效交易']}行；非交易{counts['非交易行']}行；"
            f"解析异常{counts['解析异常']}行；其他排除{counts['其他排除']}行；"
            f"原始可解析净额{raw_net:.2f}；"
            + "；".join(f"{category}净额{net:.2f}" for category, net in category_net.items())
            + f"；金额勾稽差额{diff:.2f}；原金额或方向未能解析{unknown}行未计入勾稽"
        )
    items.append(
        PrecheckItem(
            "数据入口",
            population_summary("银行流水"),
            population_summary("银行日记账"),
            "存在未完全解释的行" if population_risk else "行数去向可解释",
            "疑点" if population_risk else "通过",
            (
                "程序已处理全部可用交易；解析或行数缺口作为范围限制披露。"
                if population_risk
                else "原始数据行已按唯一类别归集；原始可解析净额包含非交易合计行，不是交易发生额。"
            ),
        )
    )

    bank_mapping_problems = _mapping_problems(bank_mapping, raw_bank.columns)
    journal_mapping_problems = _mapping_problems(journal_mapping, raw_journal.columns)
    mapping_blocked = bool(bank_mapping_problems or journal_mapping_problems)
    items.append(
        PrecheckItem(
            "必填字段",
            "完整" if not bank_mapping_problems else "；".join(bank_mapping_problems),
            "完整" if not journal_mapping_problems else "；".join(journal_mapping_problems),
            "双方完整" if not mapping_blocked else "存在缺失或无效映射",
            "无法计算" if mapping_blocked else "通过",
            "请返回选择当前金额模式所需的必填列。" if mapping_blocked else "当前金额模式所需字段完整。",
        )
    )

    bank_aux, bank_aux_low = _auxiliary_result(raw_bank, bank_mapping)
    journal_aux, journal_aux_low = _auxiliary_result(raw_journal, journal_mapping)
    aux_low = bank_aux_low or journal_aux_low
    items.append(
        PrecheckItem(
            "辅助文字完整性",
            bank_aux,
            journal_aux,
            "至少一侧偏低" if aux_low else "双方可用",
            "疑点" if aux_low else "通过",
            "摘要、对方户名等辅助文字非空率低于80%，文字匹配证据可能不足。" if aux_low else "辅助文字列可用。",
        )
    )

    return InputPrecheckReport(
        items=tuple(items),
        source_info=tuple(source_info),
        population_rows=tuple(population_rows),
        overall_control=overall_control,
        parse_errors=tuple(parse_errors),
    )
