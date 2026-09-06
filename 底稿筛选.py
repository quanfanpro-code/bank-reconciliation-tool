"""从全量核对报告按审计条件整组筛选并导出新底稿。"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import pandas as pd

from make_excel import make_excel


@dataclass(frozen=True)
class FilterCriteria:
    coverage_ratio: float | None = None
    start_date: str | None = None
    end_date: str | None = None
    include_text: tuple[str, ...] = ()
    exclude_text: tuple[str, ...] = ()
    business_types: tuple[str, ...] = ()
    statuses: tuple[str, ...] = ()
    reasons: tuple[str, ...] = ()


def _combined_text(frame: pd.DataFrame) -> pd.Series:
    if frame.empty:
        return pd.Series([], dtype=str, index=frame.index)
    columns = [column for column in frame.columns if any(key in str(column) for key in ("摘要", "对方", "依据", "原因", "类型", "检索"))]
    if not columns:
        return pd.Series("", index=frame.index, dtype=str)
    return frame[columns].fillna("").astype(str).agg("|".join, axis=1)


def _filter_groups(frame: pd.DataFrame, criteria: FilterCriteria) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    mask = pd.Series(True, index=frame.index)
    if criteria.business_types and "类型" in frame:
        mask &= frame["类型"].astype(str).isin(criteria.business_types)
    status_column = "最终状态" if "最终状态" in frame else "系统结论" if "系统结论" in frame else None
    if criteria.statuses and status_column:
        mask &= frame[status_column].astype(str).isin(criteria.statuses)
    text = _combined_text(frame)
    for keyword in criteria.include_text:
        mask &= text.str.contains(str(keyword), regex=False, na=False)
    for keyword in criteria.exclude_text:
        mask &= ~text.str.contains(str(keyword), regex=False, na=False)
    if criteria.reasons:
        reason_text = frame.get("判断依据", frame.get("处理原因", pd.Series("", index=frame.index))).astype(str)
        reason_mask = pd.Series(False, index=frame.index)
        for reason in criteria.reasons:
            reason_mask |= reason_text.str.contains(str(reason), regex=False, na=False)
        mask &= reason_mask
    date_column = "最早日期" if "最早日期" in frame else "日期" if "日期" in frame else None
    if date_column:
        dates = pd.to_datetime(frame[date_column], errors="coerce")
        if criteria.start_date:
            mask &= dates >= pd.Timestamp(criteria.start_date)
        if criteria.end_date:
            mask &= dates <= pd.Timestamp(criteria.end_date)
    selected = frame.loc[mask].copy()
    if criteria.coverage_ratio is not None and not selected.empty and "组金额" in selected:
        ratio = max(0.0, min(1.0, float(criteria.coverage_ratio)))
        ordered = selected.assign(_amount=pd.to_numeric(selected["组金额"], errors="coerce").abs().fillna(0)).sort_values(["_amount", "匹配ID"], ascending=[False, True], kind="stable")
        target = ordered["_amount"].sum() * ratio
        keep_count = int((ordered["_amount"].cumsum() < target).sum()) + (1 if target > 0 else 0)
        selected = ordered.iloc[:keep_count].drop(columns="_amount")
    return selected.reset_index(drop=True)


def export_filtered_workpaper(
    source_path: str | Path,
    output_path: str | Path,
    criteria: FilterCriteria,
    *,
    progress_callback: Callable[[float], None] | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path:
    progress = progress_callback or (lambda _value: None)
    log = log_callback or (lambda _message: None)
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output:
        raise ValueError("筛选版必须另存为新文件，不能覆盖全量报告")
    progress(0.0)
    log(f"开始读取全量报告：{source.name}")
    sheets = pd.read_excel(source, sheet_name=None)
    progress(0.2)
    log(f"全量报告读取完成：{len(sheets)} 个工作表")
    group_names = [name for name in ("逐笔匹配", "整组勾稽") if name in sheets]
    group_frames = []
    component_text: dict[str, str] = {}
    components = sheets.get("匹配组成")
    if components is not None and not components.empty and "匹配ID" in components:
        searchable_columns = [column for column in components.columns if any(key in str(column) for key in ("摘要", "对方", "辅助文字", "业务"))]
        if searchable_columns:
            component_text = (
                components.assign(_检索=components[searchable_columns].fillna("").astype(str).agg("|".join, axis=1))
                .groupby(components["匹配ID"].astype(str))["_检索"]
                .agg("|".join)
                .to_dict()
            )
    for name in group_names:
        frame = sheets[name].copy()
        frame["_来源工作表"] = name
        if "匹配ID" in frame:
            frame["_组成检索文字"] = frame["匹配ID"].astype(str).map(component_text).fillna("")
        group_frames.append(frame)
    combined_groups = pd.concat(group_frames, ignore_index=True) if group_frames else pd.DataFrame()
    selected_combined = _filter_groups(combined_groups, criteria)
    progress(0.4)
    log(f"关系筛选完成：{len(selected_combined)}/{len(combined_groups)} 组")
    filtered_groups = {}
    for name in group_names:
        selected = selected_combined.loc[selected_combined["_来源工作表"] == name].copy()
        filtered_groups[name] = selected.drop(columns=["_来源工作表", "_组成检索文字"], errors="ignore").reset_index(drop=True)
    selected_ids = {
        str(value)
        for frame in filtered_groups.values()
        if "匹配ID" in frame
        for value in frame["匹配ID"].dropna()
    }
    total_groups = combined_groups.drop(columns=["_来源工作表", "_组成检索文字"], errors="ignore")
    selected_group_frame = selected_combined.drop(columns=["_来源工作表", "_组成检索文字"], errors="ignore")
    total_amount = pd.to_numeric(total_groups.get("组金额", pd.Series(dtype=float)), errors="coerce").abs().fillna(0).sum()
    selected_amount = pd.to_numeric(selected_group_frame.get("组金额", pd.Series(dtype=float)), errors="coerce").abs().fillna(0).sum()
    explanation = pd.DataFrame(
        [
            ("全量报告", str(source)),
            ("全量关系数", len(total_groups)),
            ("筛选关系数", len(selected_group_frame)),
            ("全量关系金额", float(total_amount)),
            ("筛选关系金额", float(selected_amount)),
            ("覆盖比例", float(selected_amount / total_amount) if total_amount else 0.0),
            ("筛选条件", str(criteria)),
            ("说明", "筛选版只用于审计选项；全量核对报告保持不变。任一关系被选中时，其银行流水和序时账组成全部带出。"),
        ],
        columns=["项目", "数值"],
    )
    output_sheets: list[tuple[str, pd.DataFrame]] = [("筛选说明", explanation)]
    for name, frame in sheets.items():
        if name in filtered_groups:
            output_sheets.append((name, filtered_groups[name]))
        elif name in ("匹配组成", "其他可能对应明细") and "匹配ID" in frame:
            output_sheets.append((name, frame.loc[frame["匹配ID"].astype(str).isin(selected_ids)].copy()))
        else:
            output_sheets.append((name, frame.copy()))
    progress(0.65)
    log(f"开始写入筛选底稿：{len(output_sheets)} 个工作表")
    make_excel(output_sheets, str(output), theme="deep-navy")
    progress(1.0)
    log(f"筛选导出完成：{output}")
    return output
