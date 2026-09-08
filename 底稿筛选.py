"""从全量核对报告按审计条件整组筛选并导出新底稿。"""

from __future__ import annotations

from contextlib import closing
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from make_excel import make_excel, atomic_output_path
from 报告列识别 import identify_columns


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
    amount_basis: str = '不限'
    min_amount: Decimal | float | None = None
    max_amount: Decimal | float | None = None
    amount_absolute: bool = True


def validate_filter_criteria(criteria: FilterCriteria) -> None:
    """空结果也检查条件，避免把输错条件解释成没有需核对事项。"""
    if criteria.amount_basis not in ('不限', '银行单笔', '序时账单笔', '整组金额'):
        raise ValueError('请选择银行单笔、序时账单笔或整组金额')
    values = []
    for value in (criteria.min_amount, criteria.max_amount):
        try:
            number = None if value is None else Decimal(str(value))
        except InvalidOperation as exc:
            raise ValueError('金额必须是有效数字') from exc
        if number is not None and not number.is_finite():
            raise ValueError('金额必须是有限数字')
        if number is not None and (criteria.amount_absolute or criteria.amount_basis == '整组金额') and number < 0:
            raise ValueError('绝对金额及整组金额的上下限不能为负数')
        values.append(number)
    if all(value is not None for value in values) and values[0] > values[1]:
        raise ValueError('金额下限不能超过上限')
    if criteria.amount_basis == '不限' and any(value is not None for value in values):
        raise ValueError('填写金额上下限后，请选择金额口径')
    if criteria.coverage_ratio is not None:
        try:
            ratio = Decimal(str(criteria.coverage_ratio))
        except InvalidOperation as exc:
            raise ValueError('覆盖比例必须是0到100之间的数字') from exc
        if not ratio.is_finite() or not 0 <= ratio <= 1:
            raise ValueError('覆盖比例必须在0到100之间')
    try:
        start = date.fromisoformat(criteria.start_date) if criteria.start_date else None
        end = date.fromisoformat(criteria.end_date) if criteria.end_date else None
    except ValueError as exc:
        raise ValueError('日期必须是有效的年-月-日，例如2026-01-31') from exc
    if start and end and start > end:
        raise ValueError('开始日期不能晚于结束日期')


def _amount_mask(frame, criteria, components):
    """以原始组成命中事项，不使用其他候选或重复展示行凑金额。"""
    if criteria.amount_basis == '不限' or (criteria.min_amount is None and criteria.max_amount is None):
        return pd.Series(True, index=frame.index)
    grouped = criteria.amount_basis == '整组金额'
    data = frame if grouped else components
    key = '事项编号' if '事项编号' in frame else '匹配ID'
    column = '组金额' if grouped else '金额'
    if data is None or column not in data or (not grouped and not {key, '来源'} <= set(data)):
        raise ValueError('报告缺少金额明细或组成，无法按此口径筛选；请重新生成完整报告')
    if not grouped:
        sources = ('银行流水',) if criteria.amount_basis == '银行单笔' else ('序时账', '日记账')
        data = data.loc[data['来源'].isin(sources)]
    def matches(row):
        try:
            amount = Decimal(str(row[column]))
        except InvalidOperation as exc:
            raise ValueError('报告组成中存在无法识别的金额，已停止筛选') from exc
        if not amount.is_finite():
            raise ValueError('报告组成金额缺失，已停止筛选')
        # 旧版匹配组成将金额存为绝对值，结合收支方向恢复带符号口径。
        if not grouped and row.get('收支方向') == '支出':
            amount = -abs(amount)
        if grouped or criteria.amount_absolute:
            amount = abs(amount)
        return (criteria.min_amount is None or amount >= Decimal(str(criteria.min_amount))) and (criteria.max_amount is None or amount <= Decimal(str(criteria.max_amount)))
    hit = pd.Series([matches(row) for row in data.to_dict('records')], index=data.index, dtype=bool)
    if grouped:
        return hit
    return frame[key].astype(str).isin(set(data.loc[hit, key].astype(str)))


def _combined_text(frame: pd.DataFrame) -> pd.Series:
    if frame.empty:
        return pd.Series([], dtype=str, index=frame.index)
    columns = [column for column in frame.columns if any(key in str(column) for key in ("摘要", "对方", "依据", "原因", "类型", "检索"))]
    if not columns:
        return pd.Series("", index=frame.index, dtype=str)
    return frame[columns].fillna("").astype(str).agg("|".join, axis=1)


def _filter_groups(frame: pd.DataFrame, criteria: FilterCriteria, components: pd.DataFrame | None = None) -> pd.DataFrame:
    validate_filter_criteria(criteria)
    if frame.empty:
        return frame.copy()
    mask = _amount_mask(frame, criteria, components)
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


def _worksheet_frame(sheet, required=()) -> pd.DataFrame:
    """读取原单元格值，保留稳定事项编号，不使用可能已失效的公式缓存。"""
    columns = {position - 1: name for name, position in identify_columns(sheet, required).items()}
    return pd.DataFrame(
        [{name: row[position] for position, name in columns.items()}
         for row in sheet.iter_rows(min_row=2, values_only=True)],
        columns=list(columns.values()),
    )


def _review_groups(book) -> pd.DataFrame:
    groups = _worksheet_frame(book['复核事项索引'], ('事项编号', '程序结论', '组金额'))
    details = _worksheet_frame(book['核对明细'], ('事项编号', '来源', '金额'))
    texts = {}
    if not details.empty:
        texts = details.assign(_检索=_combined_text(details)).groupby('事项编号')['_检索'].agg('|'.join).to_dict()
    choices, remarks = {}, {}
    for name in ('人工全查', '人工抽样'):
        if name not in book:
            continue
        for row in _worksheet_frame(book[name], ('事项编号', '行别', '人工核对结果')).to_dict('records'):
            item_id = row.get('事项编号')
            if pd.isna(item_id) or row.get('行别') != '核对事项':
                continue
            choice = row.get('人工核对结果')
            choices[str(item_id)] = '' if pd.isna(choice) else str(choice).strip()
            remarks[str(item_id)] = '|'.join(str(row[column]) for column in ('核对原因', '银行记录', '序时账记录', '备注') if column in row and pd.notna(row[column]))
    groups['事项编号'] = groups['事项编号'].astype(str)
    groups['_组成检索文字'] = groups['事项编号'].map(texts).fillna('')
    groups['_人工检索文字'] = groups['事项编号'].map(remarks).fillna('')
    groups['人工核对结果'] = groups['事项编号'].map(choices).fillna('')
    # 已经选择时按该实际结果筛选，不能继续使用生成时的程序状态或旧缓存。
    groups['最终状态'] = groups['人工核对结果'].where(groups['人工核对结果'] != '', groups['程序结论'])
    if '匹配ID' not in groups:
        groups['匹配ID'] = groups['事项编号']
    else:
        groups['匹配ID'] = groups['匹配ID'].fillna(groups['事项编号'])
    return groups


def _describe_criteria(criteria: FilterCriteria) -> str:
    parts = []
    for value, label in ((criteria.start_date, '起始日期'), (criteria.end_date, '截止日期')):
        if value:
            parts.append(label + '：' + str(value))
    for values, label in ((criteria.include_text, '包含文字'), (criteria.exclude_text, '排除文字'),
                          (criteria.business_types, '业务类型'), (criteria.statuses, '核对结果'), (criteria.reasons, '判断依据')):
        if values:
            parts.append(label + '：' + '、'.join(str(value) for value in values))
    if criteria.coverage_ratio is not None:
        parts.append(f'条件内目标金额覆盖比例：{criteria.coverage_ratio:.2%}')
    if criteria.amount_basis != '不限':
        mode = '双方收支绝对金额合计较大值' if criteria.amount_basis == '整组金额' else '绝对金额' if criteria.amount_absolute else '带符号金额（收入正、支出负）'
        low = '不限' if criteria.min_amount is None else str(criteria.min_amount)
        high = '不限' if criteria.max_amount is None else str(criteria.max_amount)
        parts.append(f'{criteria.amount_basis}，{mode}，下限{low}元、上限{high}元（均含边界）；命中后保留完整事项')
    return '；'.join(parts) or '全部事项'


def _export_review_view(book, source, output, criteria, progress, log) -> Path:
    """另存原簿并隐藏非选中整组；不删除原行、不移动公式及人工填写单元格。"""
    groups = _review_groups(book)
    selected = _filter_groups(groups, criteria, _worksheet_frame(book['核对明细']))
    selected_ids = set(selected['事项编号'])
    progress(0.4)
    log(f'事项筛选完成：{len(selected)}/{len(groups)} 项')
    item_sheets = {'核对明细', '人工全查', '人工抽样', '月度差异组成', '其他对应供选择', '逐笔核对', '整组核对', '未对应记录'}
    full_sheets = {'核对结论', '月度核对', '每日统计'}
    for sheet in book:
        if sheet.title in item_sheets:
            header = identify_columns(sheet, ('事项编号',))
            if '事项编号' not in header:
                sheet.sheet_state = 'hidden'
                continue
            owner = None
            for row in range(2, sheet.max_row + 1):
                item_id = sheet.cell(row, header['事项编号']).value
                if item_id:
                    owner = str(item_id)
                elif sheet.title not in ('人工全查', '人工抽样'):
                    owner = None
                sheet.row_dimensions[row].hidden = owner not in selected_ids
                sheet.row_dimensions[row].collapsed = False
            sheet.auto_filter.filterColumn = []
            sheet.sheet_state = 'visible' if sheet.title in ('核对明细', '人工全查', '人工抽样') or any(not sheet.row_dimensions[row].hidden for row in range(2, sheet.max_row + 1)) else 'hidden'
        else:
            sheet.sheet_state = 'visible' if sheet.title in full_sheets else 'hidden'
    if '逐笔核对' in book:
        from 报告展示 import 阅读表
        for sheet in book:
            sheet.sheet_state = 'visible' if sheet.title in 阅读表 else 'hidden'
    amounts = lambda frame: pd.to_numeric(frame.get('组金额', pd.Series(dtype=float)), errors='coerce').abs().fillna(0).sum()
    total_amount, selected_amount = amounts(groups), amounts(selected)
    explanation = [
        ('项目', '数值'), ('全量报告', str(source)),
        ('全量事项数', len(groups)), ('筛选事项数', len(selected)),
        ('全量事项金额', float(total_amount)), ('筛选事项金额', float(selected_amount)),
        ('实际金额覆盖比例', float(selected_amount / total_amount) if total_amount else 0.0),
        ('筛选条件', _describe_criteria(criteria)),
        ('口径说明', '本文件为筛选视图；月度、每日及核对结论仍为全量口径。原记录、候选组成和公式全部保留，非选中行仅隐藏，用于公式计算；筛选事项金额取双方收支绝对金额合计的较大值，收付不能抵销。'),
        ('人工结果口径', '状态筛选采用人工全查、人工抽样中的实际选择；未选择时采用程序结论。人工结果和备注原样保留，打开Excel后公式自动重算。'),
    ]
    sheet = book['筛选说明'] if '筛选说明' in book else book.create_sheet('筛选说明', 0)
    for row in sheet:
        for cell in row:
            cell.value = None
    for row_number, row in enumerate(explanation, 1):
        for column, value in enumerate(row, 1):
            cell = sheet.cell(row_number, column, value)
            cell.font = Font(name='微软雅黑', size=11, bold=row_number == 1, color='FFFFFF' if row_number == 1 else '243746')
            cell.fill = PatternFill('solid', fgColor='24445B' if row_number == 1 else 'FFFFFF')
            cell.alignment = Alignment(wrap_text=True, vertical='top')
    sheet.column_dimensions['A'].width = 26
    sheet.column_dimensions['B'].width = 110
    sheet.row_dimensions[9].height = 72
    sheet.row_dimensions[10].height = 48
    sheet['B7'].number_format = '0.00%'
    sheet.freeze_panes = 'B2'
    sheet.sheet_state = 'visible'
    book.active = book.index(sheet)
    book.calculation.calcMode = 'auto'
    book.calculation.fullCalcOnLoad = True
    book.calculation.forceFullCalc = True
    progress(0.65)
    log(f'开始写入筛选底稿：保留 {len(book.sheetnames)} 个工作表及全部计算上下文')
    with atomic_output_path(output) as temporary:
        book.save(temporary)
    progress(1.0)
    log(f'筛选导出完成：{output}')
    return output


def export_filtered_workpaper(
    source_path: str | Path,
    output_path: str | Path,
    criteria: FilterCriteria,
    *,
    progress_callback: Callable[[float], None] | None = None,
    log_callback: Callable[[str], None] | None = None,
) -> Path:
    validate_filter_criteria(criteria)
    progress = progress_callback or (lambda _value: None)
    log = log_callback or (lambda _message: None)
    source = Path(source_path).resolve()
    output = Path(output_path).resolve()
    if source == output or (output.exists() and source.samefile(output)):
        raise ValueError("筛选版必须另存为新文件，不能覆盖全量报告")
    if output.exists():
        raise FileExistsError("筛选文件已存在，请另存为新文件")
    if criteria.coverage_ratio is not None and not 0 <= float(criteria.coverage_ratio) <= 1:
        raise ValueError("金额覆盖比例必须为 0 到 1 之间的有限数字")
    progress(0.0)
    log(f"开始读取全量报告：{source.name}")
    with closing(load_workbook(source, data_only=False)) as book:
        if '复核事项索引' in book and '核对明细' in book:
            progress(0.2)
            log(f'全量报告读取完成：{len(book.sheetnames)} 个工作表')
            return _export_review_view(book, source, output, criteria, progress, log)
    with closing(load_workbook(source, data_only=True)) as book:
        sheets = {sheet.title: _worksheet_frame(sheet) for sheet in book}
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
    selected_combined = _filter_groups(combined_groups, criteria, components)
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
            ("筛选条件", _describe_criteria(criteria)),
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
    with atomic_output_path(output) as temporary:
        make_excel(output_sheets, str(temporary), theme="deep-navy")
    progress(1.0)
    log(f"筛选导出完成：{output}")
    return output


def read_filter_report(source_path):
    """只读取得事项与组成，供窗口自动加载和后续条件预览复用。"""
    with closing(load_workbook(source_path, data_only=False)) as book:
        if '复核事项索引' in book and '核对明细' in book:
            groups = _review_groups(book)
            components = _worksheet_frame(book['核对明细'])
        else:
            frames = [_worksheet_frame(book[name]) for name in ('逐笔匹配', '整组勾稽') if name in book]
            if not frames:
                raise ValueError('这不是可识别的全量核对报告')
            groups = pd.concat(frames, ignore_index=True)
            components = _worksheet_frame(book['匹配组成']) if '匹配组成' in book else None
            if components is not None and '匹配ID' in components:
                texts = components.assign(_文字=_combined_text(components)).groupby('匹配ID')['_文字'].agg('|'.join)
                groups['_组成检索文字'] = groups['匹配ID'].map(texts).fillna('')
    return groups, components


def summarize_filter(groups, components, criteria):
    """对已读取的数据预览，筛选规则与正式导出一致。"""
    validate_filter_criteria(criteria)
    chosen = _filter_groups(groups, criteria, components)
    options = {}
    for key, column in [('business_types','类型'),('statuses','最终状态' if '最终状态' in groups else '系统结论')]:
        if column in groups:
            options[key] = sorted(set(groups[column].dropna().astype(str)) - {''})
    amount = pd.to_numeric(chosen.get('组金额', pd.Series(dtype=float)), errors='coerce').abs().fillna(0).sum()
    return {'total':len(groups), 'selected':len(chosen), 'amount':float(amount), 'description':_describe_criteria(criteria), 'options':options}


def preview_filter(source_path, criteria):
    """保留原只读预览入口；窗口可复用同一份读取结果。"""
    validate_filter_criteria(criteria)
    return summarize_filter(*read_filter_report(source_path), criteria)
