"""银行流水和银行日记账的自包含输入预检查。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from statistics import pstdev
from typing import Any, Iterable, Sequence

import pandas as pd


SUPPORTED_SUFFIXES = {".xlsx", ".xls", ".csv"}
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
        sheet = workbook.active
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
    title_penalty = 1 if int(raw_nonempty.max()) <= 2 else 0
    next_header_penalty = 1 if _keyword_hits(data.iloc[0].tolist()) >= 2 else 0
    stability = 1 - min(pstdev(densities), 1) if len(densities) > 1 else 1
    score = (
        keywords * 1.5
        + coverage * 2
        + (transitions / width) * 2
        + sum(densities) / len(densities)
        + stability
        - title_penalty * 4
        - next_header_penalty * 2
        - (depth - 1) * 0.1
    )
    evidence = (
        f"业务词组{keywords}个",
        f"列覆盖率{coverage:.0%}",
        f"数据区稳定度{stability:.0%}",
        f"类型转换列{transitions}个",
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
