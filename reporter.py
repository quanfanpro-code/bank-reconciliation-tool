"""
Reporter 模块 — 核对结果报表生成器

使用 make_excel deep-navy 主题输出 Excel，再通过 openpyxl 后处理添加条件格式。
"""

from typing import Optional, List, Dict, Any, Callable
from decimal import Decimal
from dataclasses import dataclass
from datetime import datetime
import json
import time
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation

from precision_engine import PrecisionEngine
from data_structures import InitialBalanceWarning, MatcherConfig
from data_loader import ParseErrorCollector
from input_precheck import InputPrecheckReport
from matcher import Matcher
from utils import round_decimal, clean_excel_string
from balance import (
    BalanceRecalculator,
    BalanceReconciler,
    check_balance_continuity as check_row_balance_continuity,
)
from make_excel import make_excel
from llm_assistant import redact_sensitive_text, sanitize_url


def _restore_numeric_cells(value: Any) -> Any:
    """将清洗后的字符串还原为数值（支持千分位文本），无法还原时原样返回。

    未达明细直接展示原始数据行，金额/余额列经 clean_excel_string 后变成字符串，
    在 Excel 中无法求和；此函数把可解析的值还原为 float，文本保持文本。
    """
    if not isinstance(value, str):
        return value
    s = value.strip().replace(',', '')
    if not s or s.lower() == 'nan':
        return value
    try:
        return float(s)
    except ValueError:
        return value


class Reporter:
    """核对结果报表生成器，使用 make_excel deep-navy 主题输出 Excel。"""

    METRICS = ['income_count', 'income_amount', 'expense_count', 'expense_amount']
    SUFFIXES = ['_bank', '_journal']
    RENAME_MAP_DAILY = {
        'date': '日期', 'income_count_bank': '银行-收入笔数', 'income_count_journal': '日记账-收入笔数',
        'income_count_diff': '收入笔数差额', 'income_amount_bank': '银行-收入金额', 'income_amount_journal': '日记账-收入金额',
        'income_amount_diff': '收入金额差额', 'expense_count_bank': '银行-支出笔数', 'expense_count_journal': '日记账-支出笔数',
        'expense_count_diff': '支出笔数差额', 'expense_amount_bank': '银行-支出金额', 'expense_amount_journal': '日记账-支出金额',
        'expense_amount_diff': '支出金额差额', 'net_bank': '银行-变动净额', 'net_journal': '日记账-变动净额',
        'balance_bank': '银行-期末余额', 'balance_journal': '日记账-期末余额', 'balance_diff': '余额差额',
    }
    RENAME_MAP_MONTHLY = {
        'month': '月份', 'income_count_bank': '银行-收入笔数', 'income_count_journal': '日记账-收入笔数',
        'income_count_diff': '收入笔数差额', 'income_amount_bank': '银行-收入金额', 'income_amount_journal': '日记账-收入金额',
        'income_amount_diff': '收入金额差额', 'expense_count_bank': '银行-支出笔数', 'expense_count_journal': '日记账-支出笔数',
        'expense_count_diff': '支出笔数差额', 'expense_amount_bank': '银行-支出金额', 'expense_amount_journal': '日记账-支出金额',
        'expense_amount_diff': '支出金额差额', 'net_bank': '银行-变动净额', 'net_journal': '日记账-变动净额',
        'balance_bank': '银行-月末余额', 'balance_journal': '日记账-月末余额', 'balance_diff': '余额差额',
    }

    def __init__(self, matcher: Matcher, raw_bank: Optional[pd.DataFrame] = None,
                 raw_journal: Optional[pd.DataFrame] = None,
                 bank_mapping: Optional[Dict[str, Any]] = None,
                  journal_mapping: Optional[Dict[str, Any]] = None,
                  logger: Optional[Callable[[str], None]] = None,
                  error_collector: Optional[ParseErrorCollector] = None,
                  precheck_report: Optional[InputPrecheckReport] = None):
        self.matcher = matcher
        self.raw_bank = raw_bank
        self.raw_journal = raw_journal
        self.bank_mapping = bank_mapping
        self.journal_mapping = journal_mapping
        self.logger = logger
        self.initial_balance_warning: Optional[InitialBalanceWarning] = None
        self.error_collector = error_collector
        self.precheck_report = precheck_report

    def _log(self, message: str) -> None:
        if self.logger:
            self.logger(message)

    @staticmethod
    def _get_ordered_columns(df: pd.DataFrame, date_col: str = 'date') -> List[str]:
        """按照指标顺序排列列名。"""
        new_columns = [date_col]
        for metric in Reporter.METRICS:
            for suffix in Reporter.SUFFIXES:
                if metric + suffix in df.columns:
                    new_columns.append(metric + suffix)
            if metric + '_diff' in df.columns:
                new_columns.append(metric + '_diff')
        for suffix in Reporter.SUFFIXES:
            if 'net' + suffix in df.columns:
                new_columns.append('net' + suffix)
        for suffix in Reporter.SUFFIXES:
            if 'balance' + suffix in df.columns:
                new_columns.append('balance' + suffix)
        if 'balance_diff' in df.columns:
            new_columns.append('balance_diff')
        return new_columns

    @staticmethod
    def _has_balance_data(df: Optional[pd.DataFrame], mapping: Optional[Dict[str, Any]] = None) -> bool:
        """判断数据中是否存在可用的余额列数据（mapping 列优先，否则按内置词表精确匹配）。"""
        if df is None or df.empty:
            return False
        col = None
        if mapping:
            cand = mapping.get('balance')
            if cand and cand in df.columns:
                col = cand
        if col is None:
            for c in df.columns:
                c_stripped, c_lower = str(c).strip(), str(c).lower().strip()
                if c_stripped in ('balance', '余额', 'std_balance') or c_lower == 'balance':
                    col = c
                    break
        if col is None:
            return False
        series = df[col]
        non_empty = (series.notna()
                     & (series.astype(str).str.strip() != '')
                     & (series.astype(str).str.strip().str.lower() != 'nan'))
        return bool(non_empty.any())

    def check_balance_continuity(self, df: pd.DataFrame, tolerance_li: int = 10,
                                 source: str = "") -> List[Dict[str, Any]]:
        """按原文件顺序复用总体控制的逐行余额连续性检查。"""
        return check_row_balance_continuity(
            df,
            tolerance_li=tolerance_li,
            source=source,
        )

    def calculate_daily_stats(self, bank: pd.DataFrame, journal: pd.DataFrame) -> pd.DataFrame:
        """计算每日统计对比数据。"""
        if bank.empty and journal.empty:
            return pd.DataFrame()
        all_dates = set(bank['date'].unique()) | set(journal['date'].unique())
        if not all_dates:
            return pd.DataFrame()

        date_range = pd.date_range(start=min(all_dates), end=max(all_dates), freq='D')

        def get_stats(df, has_bal):
            inc = df[df['amount'] > 0]
            exp = df[df['amount'] < 0]
            stats = pd.merge(
                inc.groupby('date')['amount'].agg(income_count='count', income_amount='sum'),
                exp.groupby('date')['amount'].agg(expense_count='count', expense_amount='sum'),
                on='date', how='outer'
            )
            stats = pd.merge(stats, df.groupby('date')['amount'].agg(net='sum'), on='date', how='outer')
            stats = pd.merge(pd.DataFrame({'date': date_range}), stats, on='date', how='left').fillna(0)
            stats['expense_amount'] = stats['expense_amount'].abs()
            if has_bal:
                bal = df.sort_values(['date', 'original_idx']).groupby('date')['balance'].last()
                stats = pd.merge(stats, bal, on='date', how='left')
                stats['balance'] = stats['balance'].ffill()
                stats['balance_missing'] = stats['balance'].isna()
            return stats

        b_stats = get_stats(bank, 'balance' in bank.columns and bank['balance'].notna().any())
        j_stats = get_stats(journal, 'balance' in journal.columns and journal['balance'].notna().any())

        df = pd.merge(b_stats, j_stats, on='date', how='outer', suffixes=('_bank', '_journal'))

        # 四舍五入与差额计算
        cols = [c for c in df.columns if 'amount' in c or 'net' in c or ('balance' in c and 'missing' not in c)]
        for c in cols:
            df[c] = df[c].apply(round_decimal)

        df['income_count_diff'] = df['income_count_bank'] - df['income_count_journal']
        df['income_amount_diff'] = (df['income_amount_bank'] - df['income_amount_journal']).apply(round_decimal)
        df['expense_count_diff'] = df['expense_count_bank'] - df['expense_count_journal']
        df['expense_amount_diff'] = (df['expense_amount_bank'] - df['expense_amount_journal']).apply(round_decimal)

        if 'balance_bank' in df.columns and 'balance_journal' in df.columns:
            df['balance_diff'] = (df['balance_bank'] - df['balance_journal']).apply(round_decimal)

        return df

    def calculate_monthly_stats(self, df_daily: pd.DataFrame) -> pd.DataFrame:
        """根据每日统计计算月度汇总。"""
        df = df_daily.copy()
        df['month'] = df['date'].dt.to_period('M')

        agg = {k: 'sum' for k in [
            'income_count_bank', 'income_amount_bank', 'expense_count_bank', 'expense_amount_bank', 'net_bank',
            'income_count_journal', 'income_amount_journal', 'expense_count_journal', 'expense_amount_journal', 'net_journal'
        ]}
        df_m = df.groupby('month').agg(agg).reset_index()

        last = df.sort_values('date').drop_duplicates('month', keep='last')
        bal_cols = [c for c in ['balance_bank', 'balance_journal'] if c in last.columns]
        if bal_cols:
            df_m = pd.merge(df_m, last[['month'] + bal_cols], on='month', how='left')
            for c in bal_cols:
                df_m[c] = df_m[c].ffill()

        cols = [c for c in df_m.columns if 'amount' in c or 'net' in c or 'balance' in c]
        for c in cols:
            df_m[c] = df_m[c].apply(round_decimal)

        df_m['income_count_diff'] = df_m['income_count_bank'] - df_m['income_count_journal']
        df_m['income_amount_diff'] = (df_m['income_amount_bank'] - df_m['income_amount_journal']).apply(round_decimal)
        df_m['expense_count_diff'] = df_m['expense_count_bank'] - df_m['expense_count_journal']
        df_m['expense_amount_diff'] = (df_m['expense_amount_bank'] - df_m['expense_amount_journal']).apply(round_decimal)
        if 'balance_bank' in df_m.columns and 'balance_journal' in df_m.columns:
            df_m['balance_diff'] = (df_m['balance_bank'] - df_m['balance_journal']).apply(round_decimal)

        df_m = df_m[self._get_ordered_columns(df_m, 'month')]
        # Period 类型转为字符串
        df_m['month'] = df_m['month'].astype(str)
        return df_m

    # ------------------------------------------------------------------
    # 后处理辅助方法
    # ------------------------------------------------------------------

    @staticmethod
    def _postprocess_summary(ws, has_initial_warning: bool) -> None:
        """核对结论后处理：只高亮实际的期初余额警告行。"""
        if not has_initial_warning:
            return
        warning_font = Font(color='FF0000', bold=True, size=14)
        warning_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        max_row = ws.max_row or 7
        max_col = ws.max_column or 3
        headers = {cell.value: cell.column for cell in ws[1] if cell.value is not None}
        item_column = headers.get("项目", 1)
        for r in range(2, max_row + 1):
            item_name = str(ws.cell(row=r, column=item_column).value or "")
            if "期初余额" not in item_name:
                continue
            for c in range(2, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = warning_font
                cell.fill = warning_fill

    @staticmethod
    def _postprocess_details(ws) -> None:
        """按风险等级标示差异，不把正负方向误当成好坏。"""
        max_row = ws.max_row or 1
        max_col = ws.max_column or 1

        # 查找关键列索引
        low_conf_col_idx = None
        risk_col_idx = None
        diff_col_indices = []
        amount_keywords = ['金额', '余额', '净额', '差额', '收入', '支出']

        for ci in range(2, max_col + 1):
            header = ws.cell(row=1, column=ci).value
            if header is None:
                continue
            header_str = str(header)
            header_lower = header_str.lower()
            if header_str == '_低置信度标记':
                low_conf_col_idx = ci
            if header_str == '风险等级':
                risk_col_idx = ci
            if '差额' in header_str or 'diff' in header_lower:
                diff_col_indices.append(ci)

        low_conf_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        low_conf_font = Font(color='FF0000', bold=True)
        risk_fills = {
            '低风险': PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'),
            '中风险': PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid'),
            '高风险': PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'),
            '范围未知': PatternFill(start_color='F4CCCC', end_color='F4CCCC', fill_type='solid'),
        }

        for r in range(2, max_row + 1):
            # 低置信度行着色
            is_low_conf = False
            if low_conf_col_idx:
                val = ws.cell(row=r, column=low_conf_col_idx).value
                if val is True or str(val).lower() == 'true':
                    is_low_conf = True
                    for c in range(2, max_col + 1):
                        ws.cell(row=r, column=c).fill = low_conf_fill
                        ws.cell(row=r, column=c).font = low_conf_font

            # 差额列条件格式
            for dc in diff_col_indices:
                cell = ws.cell(row=r, column=dc)
                try:
                    v = float(cell.value) if cell.value is not None else 0
                    if v != 0:
                        risk = ws.cell(row=r, column=risk_col_idx).value if risk_col_idx else '低风险'
                        cell.fill = risk_fills.get(str(risk), risk_fills['低风险'])
                except (ValueError, TypeError):
                    pass

        # 删除 _低置信度标记 列
        if low_conf_col_idx:
            ws.delete_cols(low_conf_col_idx)

    @staticmethod
    def _postprocess_diff_columns(ws) -> None:
        """统计表中的非零差额统一提示，不按正负方向判断好坏。"""
        max_row = ws.max_row or 1
        max_col = ws.max_column or 1
        difference_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

        for ci in range(2, max_col + 1):
            header = ws.cell(row=1, column=ci).value
            if header is None:
                continue
            header_str = str(header)
            if '差额' in header_str or 'diff' in header_str.lower():
                for r in range(2, max_row + 1):
                    cell = ws.cell(row=r, column=ci)
                    try:
                        v = float(cell.value) if cell.value is not None else 0
                        if v != 0:
                            cell.fill = difference_fill
                    except (ValueError, TypeError):
                        pass

    # ------------------------------------------------------------------
    # 结构化候选报告（新版本）
    # ------------------------------------------------------------------

    @staticmethod
    def _numeric(value: Any) -> Any:
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return value
        return value

    @staticmethod
    def _safe_table(frame: pd.DataFrame) -> pd.DataFrame:
        """保留数值和日期类型，清理所有外来表头和文字。"""
        safe = frame.copy()
        safe.columns = [clean_excel_string(column) for column in safe.columns]
        for column in safe.columns:
            def clean_value(value):
                if value is None:
                    return ""
                if isinstance(value, Decimal):
                    return float(value)
                if isinstance(value, (int, float, pd.Timestamp, datetime)):
                    return value
                return clean_excel_string(value)
            safe[column] = safe[column].map(clean_value)
        return safe

    def _prepare_initial_balance(self) -> tuple[bool, bool]:
        overall_control = getattr(self.matcher, "overall_control", None)
        if overall_control is not None:
            bank_initial = overall_control.bank_initial_balance or Decimal("0")
            journal_initial = overall_control.journal_initial_balance or Decimal("0")
            bank_has_balance = (
                overall_control.bank_balance_status != "未实施"
                and overall_control.bank_initial_balance is not None
            )
            journal_has_balance = (
                overall_control.journal_balance_status != "未实施"
                and overall_control.journal_initial_balance is not None
            )
        else:
            bank_initial = BalanceRecalculator.extract_initial_balance(
                self.matcher.bank,
                source_type="bank",
            )
            journal_initial = BalanceRecalculator.extract_initial_balance(
                self.matcher.journal,
                source_type="journal",
            )
            bank_has_balance = self._has_balance_data(self.matcher.bank)
            journal_has_balance = self._has_balance_data(self.matcher.journal)
        initial_diff = abs(bank_initial - journal_initial)
        balance_check_possible = bank_has_balance and journal_has_balance
        has_warning = (
            balance_check_possible and initial_diff > Decimal("0.01")
        )
        self.initial_balance_warning = InitialBalanceWarning(
            has_warning=has_warning,
            bank_initial=bank_initial,
            journal_initial=journal_initial,
            diff=initial_diff,
            message=(
                "期初余额不一致，请先核对期初余额"
                if has_warning
                else ""
            ),
        )
        return balance_check_possible, has_warning

    def _build_daily_and_monthly_tables(
        self,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        daily = self.calculate_daily_stats(
            self.matcher.bank,
            self.matcher.journal,
        )
        daily_columns = list(self.RENAME_MAP_DAILY.values())
        monthly_columns = list(self.RENAME_MAP_MONTHLY.values())
        if daily.empty:
            return (
                pd.DataFrame(columns=daily_columns),
                pd.DataFrame(columns=monthly_columns),
            )
        ordered = daily[self._get_ordered_columns(daily)]
        daily_cn = ordered.rename(
            columns={
                key: value
                for key, value in self.RENAME_MAP_DAILY.items()
                if key in ordered.columns
            }
        )
        monthly = self.calculate_monthly_stats(daily)
        monthly_cn = monthly.rename(
            columns={
                key: value
                for key, value in self.RENAME_MAP_MONTHLY.items()
                if key in monthly.columns
            }
        )
        return daily_cn, monthly_cn

    @staticmethod
    def _match_type_name(match_type: str) -> str:
        names = {
            "exact_1to1": "精确一对一",
            "tolerance_date": "日期容差",
            "amount_difference": "明显微小金额差异",
            "batch_aggregation": "批量聚合",
            "continuous_summary_group": "连续摘要整组",
            "combination_dfs": "组合求和",
            "daily_total": "日总额",
            "monthly_total": "月总额",
            "cross_month_total": "跨月多对多",
            "closed_candidate_group": "整组勾稽",
            "business_group": "业务完整组",
            "fee_net": "手续费净额",
        }
        return names.get(match_type, match_type)

    def _candidate_group_rows(self) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        candidates = list(
            getattr(self.matcher, "selected_candidates", [])
        )
        candidate_by_id = {
            item.candidate_id: item
            for item in getattr(self.matcher, "candidates", [])
        }
        for candidate in candidates:
            all_dates = candidate.bank_dates + candidate.journal_dates
            bank_total_li = (
                candidate.metrics.bank_income_li
                - candidate.metrics.bank_expense_li
            )
            journal_total_li = (
                candidate.metrics.journal_income_li
                - candidate.metrics.journal_expense_li
            )
            evidence = candidate.text_evidence
            difference_pool_ids = "；".join(
                str(value)
                for value in candidate.evidence.get(
                    "difference_pool_ids",
                    [],
                )
            )
            rows.append(
                {
                    "匹配ID": candidate.final_match_id
                    or candidate.candidate_id,
                    "候选ID": candidate.candidate_id,
                    "候选稳定键": getattr(candidate, "stable_key", ""),
                    "组成键": getattr(candidate, "composition_key", ""),
                    "阶段": candidate.match_stage,
                    "类型": self._match_type_name(candidate.match_type),
                    "最终状态": candidate.processing_status.value,
                    "风险等级": candidate.risk_level.value,
                    "系统结论": candidate.processing_status.value,
                    "判断依据": candidate.processing_reason,
                    "业务分组依据": candidate.evidence.get("business_basis", ""),
                    "批次核查线索": candidate.evidence.get("batch_review_hint", ""),
                    "其他可能对应": self._alternative_composition(candidate, candidate_by_id),
                    "竞争组其他候选总数": int(
                        candidate.evidence.get(
                            "alternative_candidate_count",
                            len(candidate.evidence.get("alternative_candidate_ids", ())),
                        )
                        or 0
                    ),
                    "本报告列示候选数": len(
                        candidate.evidence.get("alternative_candidate_ids", ())
                    ),
                    "建议动作": self._suggested_action(candidate.risk_level.value),
                    "处理原因": candidate.processing_reason,
                    "关系公式": candidate.evidence.get("relationship_formula", ""),
                    "公式差额": float(PrecisionEngine.from_integer_li(candidate.evidence.get("formula_difference_li", candidate.metrics.total_diff_li))),
                    "费用金额": float(PrecisionEngine.from_integer_li(candidate.evidence.get("fee_amount_li", 0))),
                    "综合可信度": candidate.scores.total,
                    "金额分": candidate.scores.amount,
                    "日期分": candidate.scores.date,
                    "文字分": candidate.scores.text,
                    "结构分": candidate.scores.structure,
                    "银行笔数": len(candidate.bank_idxs),
                    "日记账笔数": len(candidate.journal_idxs),
                    "银行合计": float(
                        PrecisionEngine.from_integer_li(bank_total_li)
                    ),
                    "日记账合计": float(
                        PrecisionEngine.from_integer_li(journal_total_li)
                    ),
                    "银行收入": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.bank_income_li
                        )
                    ),
                    "银行支出": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.bank_expense_li
                        )
                    ),
                    "银行净额": float(
                        PrecisionEngine.from_integer_li(bank_total_li)
                    ),
                    "日记账收入": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.journal_income_li
                        )
                    ),
                    "日记账支出": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.journal_expense_li
                        )
                    ),
                    "日记账净额": float(
                        PrecisionEngine.from_integer_li(journal_total_li)
                    ),
                    "组金额": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.group_amount_li
                        )
                    ),
                    "收入差额": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.income_diff_li
                        )
                    ),
                    "支出差额": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.expense_diff_li
                        )
                    ),
                    "总差额": float(
                        PrecisionEngine.from_integer_li(
                            candidate.metrics.total_diff_li
                        )
                    ),
                    "银_日期": min(candidate.bank_dates)
                    if candidate.bank_dates
                    else "",
                    "账_日期": min(candidate.journal_dates)
                    if candidate.journal_dates
                    else "",
                    "最早日期": min(all_dates) if all_dates else "",
                    "最晚日期": max(all_dates) if all_dates else "",
                    "文字支持": "；".join(
                        evidence.supporting_fields if evidence else ()
                    ),
                    "文字冲突": "；".join(
                        evidence.conflicting_fields if evidence else ()
                    ),
                    "大模型判断": (
                        candidate.llm_decision.reason
                        if candidate.llm_decision
                        else ""
                    ),
                    "差异池ID": difference_pool_ids,
                    "是否使用大模型": (
                        "是" if candidate.llm_decision else "否"
                    ),
                    "纳入风险池": (
                        "是"
                        if candidate.evidence.get(
                            "included_in_risk_pool",
                            False,
                        )
                        else "否"
                    ),
                }
            )

        if rows:
            return rows

        bank = self.matcher.bank
        journal = self.matcher.journal
        ids = sorted(
            {
                value
                for value in list(bank["match_id"]) + list(journal["match_id"])
                if value
            },
            key=str,
        )
        for match_id in ids:
            bank_group = bank[bank["match_id"] == match_id]
            journal_group = journal[journal["match_id"] == match_id]
            dates = list(bank_group["date"]) + list(journal_group["date"])
            match_type = (
                bank_group.iloc[0]["match_type"]
                if not bank_group.empty
                else journal_group.iloc[0]["match_type"]
            )
            confidence = (
                bank_group.iloc[0]["confidence"]
                if not bank_group.empty
                else journal_group.iloc[0]["confidence"]
            )
            rows.append(
                {
                    "匹配ID": str(match_id),
                    "候选ID": "",
                    "阶段": "",
                    "类型": self._match_type_name(str(match_type)),
                    "最终状态": "自动确认",
                    "风险等级": "正常",
                    "系统结论": "自动确认",
                    "判断依据": "兼容旧匹配记录",
                    "建议动作": "无需额外处理",
                    "处理原因": "兼容旧匹配记录",
                    "综合可信度": {"高": 90, "中": 75, "低": 60}.get(
                        confidence,
                        0,
                    ),
                    "金额分": "",
                    "日期分": "",
                    "文字分": "",
                    "结构分": "",
                    "银行笔数": len(bank_group),
                    "日记账笔数": len(journal_group),
                    "银行合计": float(bank_group["amount"].sum()),
                    "日记账合计": float(journal_group["amount"].sum()),
                    "银行收入": float(
                        bank_group.loc[
                            bank_group["amount"] > 0,
                            "amount",
                        ].sum()
                    ),
                    "银行支出": float(
                        -bank_group.loc[
                            bank_group["amount"] < 0,
                            "amount",
                        ].sum()
                    ),
                    "银行净额": float(bank_group["amount"].sum()),
                    "日记账收入": float(
                        journal_group.loc[
                            journal_group["amount"] > 0,
                            "amount",
                        ].sum()
                    ),
                    "日记账支出": float(
                        -journal_group.loc[
                            journal_group["amount"] < 0,
                            "amount",
                        ].sum()
                    ),
                    "日记账净额": float(journal_group["amount"].sum()),
                    "组金额": max(
                        float(bank_group["amount"].abs().sum()),
                        float(journal_group["amount"].abs().sum()),
                    ),
                    "收入差额": "",
                    "支出差额": "",
                    "总差额": abs(
                        float(bank_group["amount"].sum())
                        - float(journal_group["amount"].sum())
                    ),
                    "银_日期": bank_group["date"].min()
                    if not bank_group.empty
                    else "",
                    "账_日期": journal_group["date"].min()
                    if not journal_group.empty
                    else "",
                    "最早日期": min(dates) if dates else "",
                    "最晚日期": max(dates) if dates else "",
                    "文字支持": "",
                    "文字冲突": "",
                    "大模型判断": "",
                    "差异池ID": "",
                    "是否使用大模型": "否",
                    "纳入风险池": "否",
                }
            )
        return rows

    def _build_match_group_table(self) -> pd.DataFrame:
        columns = [
            "系统结论", "风险等级", "判断依据", "业务分组依据", "批次核查线索", "其他可能对应",
            "竞争组其他候选总数", "本报告列示候选数", "建议动作", "匹配ID",
            "候选稳定键", "组成键", "类型", "银行笔数", "日记账笔数", "银行合计", "日记账合计",
            "总差额", "最早日期", "最晚日期", "候选ID", "阶段", "最终状态", "处理原因",
            "综合可信度", "金额分", "日期分", "文字分", "结构分",
            "银行收入", "银行支出", "银行净额", "日记账收入",
            "日记账支出", "日记账净额", "组金额", "收入差额",
            "支出差额", "银_日期", "账_日期", "文字支持", "文字冲突",
            "关系公式", "公式差额", "费用金额",
            "大模型判断", "差异池ID", "是否使用大模型",
            "纳入风险池",
        ]
        rows = self._candidate_group_rows()
        rows.sort(
            key=lambda row: (
                row.get("最早日期") or pd.Timestamp.max,
                row.get("组金额", 0),
                str(row.get("匹配ID", "")),
            )
        )
        return pd.DataFrame(rows, columns=columns)

    def _alternative_composition(self, candidate: Any, candidate_by_id: dict) -> str:
        """展示仍有可能的对应关系，行号直接指向两份原始文件。"""
        candidate_ids = candidate.evidence.get("alternative_candidate_ids", ())
        total = int(
            candidate.evidence.get("alternative_candidate_count", len(candidate_ids))
            or 0
        )
        if total <= len(candidate_ids) and len(candidate_ids) > 10:
            return f"另有{total}套可能对应；完整组成见“其他可能对应明细”，按匹配ID查找。"
        descriptions = []
        if total > len(candidate_ids):
            descriptions.append(
                f"所在竞争组除本关系外共{total}套候选关系；"
                f"本报告按稳定优先顺序仅列{len(candidate_ids)}套，"
                f"其余{total - len(candidate_ids)}套未逐项展开。"
            )
        for candidate_id in candidate_ids:
            other = candidate_by_id.get(candidate_id)
            if other is None:
                continue
            sides = []
            for label, frame, indexes in (
                ("银行", self.matcher.bank, other.bank_idxs),
                ("序时账", self.matcher.journal, other.journal_idxs),
            ):
                row_numbers = (
                    "、".join(str(int(frame.loc[index].get("original_file_row", frame.loc[index].get("original_idx", index)))) for index in indexes)
                    if len(indexes) <= 30 else f"共{len(indexes)}笔，逐行见“其他可能对应明细”"
                )
                sides.append(f"{label}原文件行：{row_numbers}")
            descriptions.append(
                f"对应候选：{candidate_id}；" + "；".join(sides)
            )
        return "\n".join(descriptions)

    def _build_alternative_component_table(self) -> pd.DataFrame:
        """每套竞争关系逐笔列出，不受一个Excel单元格长度限制。"""
        candidates = {c.candidate_id: c for c in getattr(self.matcher, "candidates", [])}
        rows = []
        for selected in getattr(self.matcher, "selected_candidates", []):
            for number, candidate_id in enumerate(selected.evidence.get("alternative_candidate_ids", ()), 1):
                other = candidates.get(candidate_id)
                if other is None:
                    continue
                for source, frame, indexes in (
                    ("银行流水", self.matcher.bank, other.bank_idxs),
                    ("银行存款序时账", self.matcher.journal, other.journal_idxs),
                ):
                    for index in indexes:
                        row = frame.loc[index]
                        amount = float(row["amount"])
                        rows.append({
                            "匹配ID": selected.final_match_id or selected.candidate_id,
                            "竞争组其他候选总数": int(
                                selected.evidence.get(
                                    "alternative_candidate_count",
                                    len(selected.evidence.get("alternative_candidate_ids", ())),
                                )
                                or 0
                            ),
                            "本报告列示候选数": len(
                                selected.evidence.get("alternative_candidate_ids", ())
                            ),
                            "是否截断列示": (
                                "是"
                                if selected.evidence.get("alternative_candidates_truncated")
                                else "否"
                            ),
                            "候选稳定键": getattr(other, "stable_key", ""),
                            "组成键": getattr(other, "composition_key", ""),
                            "对应候选ID": candidate_id,
                            "对应方案": number,
                            "对应性质": "竞争关系，尚不能唯一确认",
                            "来源": source,
                            "原文件行号": int(row.get("original_file_row", row.get("original_idx", index))),
                            "日期": row.get("date", ""),
                            "金额": abs(amount),
                            "收支方向": "收入" if amount > 0 else "支出" if amount < 0 else "零金额",
                            "摘要": row.get("summary", ""),
                            "辅助文字": self._auxiliary_text(row),
                            "业务分组依据": other.evidence.get("business_basis", ""),
                            **self._source_evidence_columns(row),
                        })
        return pd.DataFrame(rows)

    @staticmethod
    def _auxiliary_text(row: pd.Series) -> str:
        fields = row.get("aux_text_fields", {})
        if isinstance(fields, dict):
            return "；".join(
                f"{key}：{value}"
                for key, value in sorted(fields.items())
                if value is not None and str(value).strip()
            )
        return ""

    @staticmethod
    def _source_evidence_columns(row: pd.Series) -> Dict[str, Any]:
        """把金额和凭证证据展开成审计人员可直接查看的普通列。"""
        def mapping_text(value: Any, *, exclude: Any = None) -> str:
            if not isinstance(value, dict):
                return ""
            parts = []
            for key in sorted(value, key=lambda item: str(item)):
                if exclude is not None and str(key) == str(exclude):
                    continue
                item = value[key]
                if item is None or (isinstance(item, str) and not item.strip()):
                    continue
                parts.append(
                    f"{clean_excel_string(key)}={clean_excel_string(item)}"
                )
            return "；".join(parts)

        amount_evidence = row.get("amount_evidence", {})
        if not isinstance(amount_evidence, dict):
            amount_evidence = {}
        date_evidence = row.get("date_evidence", {})
        if not isinstance(date_evidence, dict):
            date_evidence = {}
        scope_evidence = row.get("scope_evidence", {})
        if not isinstance(scope_evidence, dict):
            scope_evidence = {}
        voucher_evidence = row.get("voucher_evidence", {})
        if not isinstance(voucher_evidence, dict):
            voucher_evidence = {}
        amount = row.get("amount", None)
        try:
            standardized_amount = float(amount) if amount is not None else None
        except (TypeError, ValueError):
            standardized_amount = amount
        inherited = voucher_evidence.get("inherited")
        return {
            "原日期列名": date_evidence.get("date_column", ""),
            "原日期值": date_evidence.get("original_value"),
            "其他业务日期列及原值": mapping_text(
                date_evidence.get("related_date_values", {}),
                exclude=date_evidence.get("date_column"),
            ),
            "采用金额口径": amount_evidence.get("amount_basis", ""),
            "原始金额模式": amount_evidence.get("mode", ""),
            "原借方列名": amount_evidence.get("debit_column", ""),
            "原借方值": amount_evidence.get("debit_value"),
            "原贷方列名": amount_evidence.get("credit_column", ""),
            "原贷方值": amount_evidence.get("credit_value"),
            "原金额列名": amount_evidence.get("amount_column", ""),
            "原金额值": amount_evidence.get("amount_value"),
            "原方向列名": amount_evidence.get("direction_column", ""),
            "原方向值": amount_evidence.get("direction_value"),
            "原始相关金额列及原值": mapping_text(
                amount_evidence.get("related_amount_values", {})
            ),
            "标准化净额": standardized_amount,
            "原账户列名": scope_evidence.get("account_column", ""),
            "原账户值": scope_evidence.get("account_value", ""),
            "原币种列名": scope_evidence.get("currency_column", ""),
            "原币种值": scope_evidence.get("currency_value", ""),
            "其他币种列及原值": mapping_text(
                scope_evidence.get("related_currency_values", {}),
                exclude=scope_evidence.get("currency_column"),
            ),
            "原凭证字列名": voucher_evidence.get("voucher_word_column", ""),
            "原凭证字值": voucher_evidence.get("voucher_word_original_value"),
            "解析后凭证字": voucher_evidence.get("voucher_word_resolved_value"),
            "原凭证列名": voucher_evidence.get("voucher_column", ""),
            "原凭证值": voucher_evidence.get("original_value"),
            "解析后凭证号": voucher_evidence.get("resolved_value"),
            "凭证是否继承": (
                "是" if inherited is True else "否" if inherited is False else ""
            ),
        }

    def _build_match_component_table(self) -> pd.DataFrame:
        columns = [
            "匹配ID", "候选ID", "来源", "原文件行号", "日期", "金额", "收支方向",
            "摘要", "辅助文字", "凭证字", "凭证号", "类型", "处理状态",
            "纳入风险池", "候选稳定键", "组成键",
            "原日期列名", "原日期值", "其他业务日期列及原值",
            "采用金额口径", "原始金额模式", "原借方列名", "原借方值",
            "原贷方列名", "原贷方值", "原金额列名", "原金额值",
            "原方向列名", "原方向值", "原始相关金额列及原值", "标准化净额",
            "原账户列名", "原账户值", "原币种列名", "原币种值",
            "其他币种列及原值", "原凭证字列名",
            "原凭证字值", "解析后凭证字", "原凭证列名",
            "原凭证值", "解析后凭证号", "凭证是否继承",
        ]
        rows: list[dict[str, Any]] = []
        candidates = list(
            getattr(self.matcher, "selected_candidates", [])
        )
        if candidates:
            for candidate in sorted(
                candidates,
                key=lambda item: (
                    min(item.bank_dates + item.journal_dates)
                    if item.bank_dates or item.journal_dates
                    else pd.Timestamp.max,
                    item.final_match_id,
                ),
            ):
                for source_name, frame, indexes in (
                    ("银行流水", self.matcher.bank, candidate.bank_idxs),
                    ("日记账", self.matcher.journal, candidate.journal_idxs),
                ):
                    for index in indexes:
                        row = frame.loc[index]
                        amount = float(row.get("amount", 0))
                        rows.append(
                            {
                                "匹配ID": candidate.final_match_id,
                                "候选ID": candidate.candidate_id,
                                "候选稳定键": getattr(candidate, "stable_key", ""),
                                "组成键": getattr(candidate, "composition_key", ""),
                                "来源": source_name,
                                "原文件行号": int(
                                    row.get(
                                        "original_file_row",
                                        row.get("original_idx", index),
                                    )
                                ),
                                "日期": row.get("date", ""),
                                "金额": abs(amount),
                                "收支方向": (
                                    "收入" if amount > 0 else "支出" if amount < 0 else "零金额"
                                ),
                                "摘要": row.get("summary", ""),
                                "辅助文字": self._auxiliary_text(row),
                                "凭证字": (
                                    row.get("voucher_word", "")
                                    if source_name == "日记账"
                                    else ""
                                ),
                                "凭证号": (
                                    row.get("voucher_no", "")
                                    if source_name == "日记账"
                                    else ""
                                ),
                                "类型": self._match_type_name(
                                    candidate.match_type
                                ),
                                "处理状态": candidate.processing_status.value,
                                "纳入风险池": (
                                    "是"
                                    if candidate.evidence.get(
                                        "included_in_risk_pool",
                                        False,
                                    )
                                    else "否"
                                ),
                                **self._source_evidence_columns(row),
                            }
                        )
        else:
            for source_name, frame in (
                ("银行流水", self.matcher.bank),
                ("日记账", self.matcher.journal),
            ):
                matched = frame[frame["matched"]].sort_values(
                    ["date", "amount_decimal"],
                )
                for index, row in matched.iterrows():
                    amount = float(row.get("amount", 0))
                    rows.append(
                        {
                            "匹配ID": row.get("match_id", ""),
                            "候选ID": "",
                            "来源": source_name,
                            "原文件行号": int(
                                row.get(
                                    "original_file_row",
                                    row.get("original_idx", index),
                                )
                            ),
                            "日期": row.get("date", ""),
                            "金额": abs(amount),
                            "收支方向": (
                                "收入" if amount > 0 else "支出" if amount < 0 else "零金额"
                            ),
                            "摘要": row.get("summary", ""),
                            "辅助文字": self._auxiliary_text(row),
                            "凭证字": (
                                row.get("voucher_word", "")
                                if source_name == "日记账"
                                else ""
                            ),
                            "凭证号": (
                                row.get("voucher_no", "")
                                if source_name == "日记账"
                                else ""
                            ),
                            "类型": self._match_type_name(
                                str(row.get("match_type", ""))
                            ),
                            "处理状态": "自动确认",
                            "纳入风险池": "否",
                            "候选稳定键": "",
                            "组成键": "",
                            **self._source_evidence_columns(row),
                        }
                    )
        return pd.DataFrame(rows, columns=columns)

    def _build_trivial_table(self) -> pd.DataFrame:
        columns = [
            "系统结论", "风险等级", "判断依据", "建议动作", "月份",
            "差异池", "池累计金额", "池是否超限", "匹配候选ID",
            "差异金额", "纳入风险池", "处理状态",
        ]
        rows = []
        for pool in getattr(self.matcher, "difference_pools", []):
            for component in pool.components:
                rows.append(
                    {
                        "系统结论": pool.processing_status.value,
                        "风险等级": pool.risk_level.value,
                        "判断依据": pool.processing_reason,
                        "建议动作": self._suggested_action(pool.risk_level.value),
                        "月份": pool.month,
                        "差异池": pool.pool_type.value,
                        "池累计金额": float(
                            PrecisionEngine.from_integer_li(
                                pool.total_diff_li
                            )
                        ),
                        "池是否超限": (
                            "是"
                            if pool.exceeds_performance_materiality
                            else "否"
                        ),
                        "匹配候选ID": component.candidate_id,
                        "差异金额": float(
                            PrecisionEngine.from_integer_li(
                                component.diff_li
                            )
                        ),
                        "纳入风险池": (
                            "是"
                            if component.included_in_risk_pool
                            else "否"
                        ),
                        "处理状态": pool.processing_status.value,
                    }
                )
        return pd.DataFrame(rows, columns=columns)

    def _source_snapshot(
        self,
        frame: pd.DataFrame,
        indexes: List[int],
        prefix: str,
    ) -> Dict[str, Any]:
        """把一侧多笔组成压缩成可直接在复核表查看的证据。"""
        valid_indexes = sorted(
            {
                int(index)
                for index in indexes
                if int(index) in frame.index
            }
        )
        if not valid_indexes:
            return {
                f"{prefix}原文件行号": "",
                f"{prefix}日期": "",
                f"{prefix}金额": 0.0,
                f"{prefix}逐笔金额": "",
                f"{prefix}摘要": "",
                f"{prefix}辅助文字": "",
            }
        selected = frame.loc[valid_indexes]
        return {
            f"{prefix}原文件行号": "；".join(
                str(
                    int(
                        row.get(
                            "original_file_row",
                            row.get("original_idx", index),
                        )
                    )
                )
                for index, row in selected.iterrows()
            ),
            f"{prefix}日期": "；".join(
                pd.Timestamp(value).strftime("%Y-%m-%d")
                for value in selected["date"]
            ),
            f"{prefix}金额": float(selected["amount"].sum()),
            f"{prefix}逐笔金额": "；".join(
                str(float(value)) for value in selected["amount"]
            ),
            f"{prefix}摘要": "；".join(
                str(value)
                for value in selected["summary"]
                if str(value).strip()
            ),
            f"{prefix}辅助文字": "；".join(
                self._auxiliary_text(row)
                for _, row in selected.iterrows()
                if self._auxiliary_text(row)
            ),
        }

    def _candidate_snapshots(
        self,
        candidate,
    ) -> Dict[str, Any]:
        snapshots = self._source_snapshot(
            self.matcher.bank,
            list(candidate.bank_idxs),
            "银行",
        )
        snapshots.update(
            self._source_snapshot(
                self.matcher.journal,
                list(candidate.journal_idxs),
                "日记账",
            )
        )
        return snapshots

    def _build_unmatched_table(self, source: str) -> pd.DataFrame:
        frame = (
            self.matcher.bank
            if source == "bank"
            else self.matcher.journal
        )
        raw = self.raw_bank if source == "bank" else self.raw_journal
        unmatched = frame[~frame["matched"]]
        evidence_columns = list(
            self._source_evidence_columns(pd.Series(dtype="object")).keys()
        )
        if raw is not None and not unmatched.empty:
            valid = unmatched[unmatched["original_idx"].map(lambda index: 0 <= int(index) - 1 < len(raw))]
            positions = [int(index) - 1 for index in valid["original_idx"]]
            result = raw.iloc[positions].copy()
            if "原文件行号" in result.columns:
                label = "输入表的原文件行号"
                while label in result.columns:
                    label = "输入表的" + label
                result = result.rename(columns={"原文件行号": label})
            result.insert(0, "原文件行号", valid.get("original_file_row", valid["original_idx"]).map(int).tolist())
            evidence_rows = [
                self._source_evidence_columns(row)
                for _, row in valid.iterrows()
            ]
            for column in evidence_columns:
                result[column] = [item[column] for item in evidence_rows]
            for column in result.columns:
                text = str(column).lower()
                if any(
                    keyword in text
                    for keyword in (
                        "金额", "amount", "发生额", "余额", "balance"
                    )
                ):
                    result[column] = result[column].map(
                        _restore_numeric_cells
                    )
            return result
        columns = ["日期", "金额", "摘要", "原文件行号", *evidence_columns]
        if unmatched.empty:
            return pd.DataFrame(columns=columns)
        result = pd.DataFrame(
            {
                "日期": unmatched["date"],
                "金额": unmatched["amount"].map(float),
                "摘要": unmatched["summary"],
                "原文件行号": unmatched.get(
                    "original_file_row",
                    unmatched.get("original_idx", unmatched.index),
                ),
            }
        )
        evidence_rows = [
            self._source_evidence_columns(row)
            for _, row in unmatched.iterrows()
        ]
        for column in evidence_columns:
            result[column] = [item[column] for item in evidence_rows]
        return result.sort_values(
            ["日期", "金额"],
            kind="stable",
        )

    def _build_parameter_table(
        self,
        config: MatcherConfig,
        date_format: str,
    ) -> pd.DataFrame:
        rows = [
            ("实际执行重要性水平", float(config.performance_materiality)),
            (
                "明显微小错报临界值",
                float(config.clearly_trivial_threshold),
            ),
            ("自动确认最低综合可信度", config.auto_confirm_score),
            ("日期容差天数", config.tolerance_days),
            ("组合窗口天数", config.dfs_date_window),
            ("组合最大深度", config.max_dfs_depth),
            ("批量最少笔数", config.batch_min_count),
            ("最大候选数", config.max_candidates),
            (
                "组合搜索每来源节点上限",
                config.combination_node_limit_per_source,
            ),
            (
                "组合搜索单任务时间上限（秒）",
                config.combination_task_timeout_seconds,
            ),
            (
                "组合搜索全局时间上限（秒）",
                config.combination_global_time_limit_seconds,
            ),
            ("是否允许异号", "是" if config.allow_mixed_sign else "否"),
            ("日期格式", date_format),
            (
                "请求随机种子",
                getattr(self.matcher, "run_parameters", {}).get(
                    "requested_random_seed",
                    config.random_seed,
                ),
            ),
            (
                "实际随机种子",
                getattr(self.matcher, "run_parameters", {}).get(
                    "actual_random_seed",
                    config.random_seed,
                ),
            ),
            ("运行时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ]
        search = getattr(self.matcher, "run_parameters", {}).get("combination_search", {})
        for key, label in (
            ("business_group_bank_rows", "完整业务组覆盖银行笔数"),
            ("business_group_journal_rows", "完整业务组覆盖序时账笔数"),
            ("generic_source_rows", "通用组合搜索来源笔数"),
            ("fully_searched_source_rows", "组合搜索预算内完整来源笔数"),
            ("truncated_source_rows", "组合候选发生截断的来源笔数"),
            ("depth_limited_source_rows", "组合搜索深度受限来源笔数"),
            ("node_budget_exhausted_source_rows", "组合搜索节点预算耗尽来源笔数"),
            ("task_timeout_source_rows", "组合搜索单任务超时来源笔数"),
            ("global_timeout_unprocessed_source_rows", "组合搜索全局超时未处理来源笔数"),
            ("worker_failure_source_rows", "组合搜索工作进程故障来源笔数"),
            ("budget_exhausted_source_rows", "组合搜索预算耗尽来源笔数"),
            ("candidate_limit", "通用组合实际候选上限"),
        ):
            if key in search:
                rows.append((label, search[key]))
        if "search_budget" in search:
            rows.append((
                "组合搜索预算",
                json.dumps(search["search_budget"], ensure_ascii=False, sort_keys=True),
            ))
        candidate_search = getattr(
            self.matcher,
            "run_parameters",
            {},
        ).get("candidate_search", {})
        stage_labels = {
            "business_group": "业务完整组",
            "whitelist": "白名单",
            "exact": "精确",
            "tolerance": "容差",
            "atomic_voucher": "完整凭证",
            "generic_combination": "通用组合",
        }
        metric_labels = {
            "examined": "候选检查数",
            "retained": "候选保留数",
            "truncated_source_rows": "候选截断来源笔数",
        }
        for stage, stage_label in stage_labels.items():
            stage_stats = candidate_search.get(stage, {})
            for metric, metric_label in metric_labels.items():
                if metric in stage_stats:
                    rows.append(
                        (
                            f"{stage_label}{metric_label}",
                            stage_stats[metric],
                        )
                    )
        selection = getattr(self.matcher, "run_parameters", {}).get(
            "selection_optimization", {}
        )
        for key, label in (
            ("component_count", "竞争组数量"),
            ("exact_components", "整体最优完整求解组数"),
            ("fallback_components", "确定性降级组数"),
            ("largest_component_candidates", "最大竞争组候选数"),
            ("exact_component_limit", "整体最优候选数上限"),
            ("stopped", "整体求解是否收到中止信号"),
            ("search_fully_exhausted", "整体求解是否全部穷尽"),
        ):
            if key in selection:
                value = selection[key]
                if isinstance(value, bool):
                    value = "是" if value else "否"
                rows.append((label, value))
        assistant = getattr(self.matcher, "llm_assistant", None)
        assistant_config = getattr(assistant, "config", None)
        rows.append(
            (
                "大模型辅助",
                "启用"
                if assistant_config is not None
                and getattr(assistant_config, "enabled", False)
                else "关闭",
            )
        )
        if assistant_config is not None and getattr(
            assistant_config,
            "enabled",
            False,
        ):
            rows.extend(
                [
                    ("大模型模式", getattr(assistant_config, "mode", "")),
                    ("大模型协议", getattr(assistant_config, "protocol", "")),
                    ("大模型模型", getattr(assistant_config, "model", "")),
                    (
                        "大模型服务地址",
                        sanitize_url(
                            getattr(assistant_config, "base_url", "")
                        ),
                    ),
                ]
            )
        return pd.DataFrame(rows, columns=["参数名称", "参数值"])

    def _build_llm_table(self) -> pd.DataFrame:
        columns = [
            "请求ID", "候选ID", "是否模型选择", "银行原文件行号",
            "日记账原文件行号", "银行日期", "日记账日期", "银行金额",
            "日记账金额", "银行摘要", "日记账摘要", "银行辅助文字",
            "日记账辅助文字", "实际发送字段", "服务", "调用协议",
            "接口地址", "模型", "本地综合可信度", "本地文字分",
            "金额分", "日期分", "文字分", "结构分", "模型语义分",
            "最终综合可信度", "判断理由", "支持证据", "冲突证据",
            "不确定性", "建议状态", "最终状态", "组金额", "差异金额",
            "实际执行重要性水平", "明显微小错报临界值",
            "自动确认最低综合可信度", "开始时间", "耗时毫秒", "用量",
            "是否降级", "错误", "脱敏原始回答",
        ]
        candidate_by_id = {
            candidate.candidate_id: candidate
            for candidate in getattr(self.matcher, "candidates", [])
        }
        selected_ids = {
            candidate.candidate_id
            for candidate in getattr(
                self.matcher,
                "selected_candidates",
                [],
            )
        }
        assistant_config = getattr(
            getattr(self.matcher, "llm_assistant", None),
            "config",
            None,
        )
        endpoint = sanitize_url(
            getattr(assistant_config, "base_url", "")
        )
        config = self.matcher.config
        rows = []
        for record in getattr(self.matcher, "llm_records", []):
            candidate_ids = tuple(record.candidate_ids)
            if not candidate_ids and record.selected_candidate_id:
                candidate_ids = (record.selected_candidate_id,)
            if not candidate_ids:
                candidate_ids = ("",)
            for candidate_id in candidate_ids:
                candidate = candidate_by_id.get(candidate_id)
                row = {
                    "请求ID": record.request_id,
                    "候选ID": candidate_id,
                    "是否模型选择": (
                        "是"
                        if candidate_id
                        and candidate_id
                        == record.selected_candidate_id
                        else "否"
                    ),
                    "实际发送字段": "；".join(record.sent_fields),
                    "服务": record.provider,
                    "调用协议": record.protocol,
                    "接口地址": endpoint,
                    "模型": record.model,
                    "本地综合可信度": (
                        candidate.evidence.get(
                            "pre_llm_total_score",
                            candidate.scores.total,
                        )
                        if candidate is not None
                        else ""
                    ),
                    "本地文字分": (
                        candidate.text_evidence.local_score
                        if candidate is not None
                        and candidate.text_evidence is not None
                        else ""
                    ),
                    "金额分": (
                        candidate.scores.amount
                        if candidate is not None
                        else ""
                    ),
                    "日期分": (
                        candidate.scores.date
                        if candidate is not None
                        else ""
                    ),
                    "文字分": (
                        candidate.scores.text
                        if candidate is not None
                        else ""
                    ),
                    "结构分": (
                        candidate.scores.structure
                        if candidate is not None
                        else ""
                    ),
                    "模型语义分": record.semantic_score,
                    "最终综合可信度": (
                        candidate.scores.total
                        if candidate is not None
                        else ""
                    ),
                    "判断理由": redact_sensitive_text(record.reason),
                    "支持证据": redact_sensitive_text(
                        "；".join(record.supporting_evidence)
                    ),
                    "冲突证据": redact_sensitive_text(
                        "；".join(record.conflicting_evidence)
                    ),
                    "不确定性": record.uncertainty,
                    "建议状态": record.suggested_status,
                    "最终状态": (
                        candidate.processing_status.value
                        if candidate is not None
                        and candidate_id in selected_ids
                        else ""
                    ),
                    "组金额": (
                        float(
                            PrecisionEngine.from_integer_li(
                                candidate.metrics.group_amount_li
                            )
                        )
                        if candidate is not None
                        else ""
                    ),
                    "差异金额": (
                        float(
                            PrecisionEngine.from_integer_li(
                                candidate.metrics.total_diff_li
                            )
                        )
                        if candidate is not None
                        else ""
                    ),
                    "实际执行重要性水平": float(
                        config.performance_materiality
                    ),
                    "明显微小错报临界值": float(
                        config.clearly_trivial_threshold
                    ),
                    "自动确认最低综合可信度": (
                        config.auto_confirm_score
                    ),
                    "开始时间": record.started_at,
                    "耗时毫秒": record.duration_ms,
                    "用量": json.dumps(
                        record.usage,
                        ensure_ascii=False,
                        sort_keys=True,
                    ),
                    "是否降级": "是" if record.fallback_used else "否",
                    "错误": redact_sensitive_text(record.error),
                    "脱敏原始回答": redact_sensitive_text(
                        record.raw_response
                    ),
                }
                if candidate is not None:
                    row.update(self._candidate_snapshots(candidate))
                rows.append(row)
        return pd.DataFrame(rows, columns=columns)

    def _build_balance_tables(
        self,
        *,
        bank_has_balance: bool,
        journal_has_balance: bool,
    ) -> dict[str, pd.DataFrame]:
        tables: dict[str, pd.DataFrame] = {}
        if bank_has_balance and journal_has_balance:
            bank_balances = BalanceRecalculator().recalculate(
                self.matcher.bank
            )
            journal_balances = BalanceRecalculator().recalculate(
                self.matcher.journal
            )
            differences = BalanceReconciler(
                bank_balances=bank_balances,
                journal_balances=journal_balances,
            ).generate_diff_report()
            if differences:
                tables["余额差异明细"] = pd.DataFrame(
                    [
                        {
                            "日期": item.date,
                            "银行余额": float(item.bank_balance or 0),
                            "日记账余额": float(item.journal_balance or 0),
                            "差额": float(item.diff or 0),
                            "差异类型": item.diff_type,
                        }
                        for item in differences
                    ]
                )
        overall_control = getattr(self.matcher, "overall_control", None)
        if overall_control is not None:
            continuity_rows = list(overall_control.continuity_anomalies)
        else:
            continuity_rows = []
            if bank_has_balance:
                continuity_rows.extend(
                    self.check_balance_continuity(
                        self.matcher.bank,
                        source="银行流水",
                    )
                )
            if journal_has_balance:
                continuity_rows.extend(
                    self.check_balance_continuity(
                        self.matcher.journal,
                        source="日记账",
                    )
                )
        if continuity_rows:
            tables["余额连续性异常"] = pd.DataFrame(continuity_rows)
        return tables

    @staticmethod
    def _suggested_action(risk_level: str) -> str:
        return {
            "正常": "无需额外处理",
            "低风险": "随同常规核对留存",
            "中风险": "结合摘要和期间优先核查",
            "高风险": "优先核查业务依据及必要的账务处理",
            "范围未知": "先关注核对范围，再使用明细结论",
        }.get(risk_level, "结合明细关注异常原因")

    def _build_business_summary_table(
        self,
        config: MatcherConfig,
        *,
        balance_check_possible: bool,
    ) -> pd.DataFrame:
        candidates = list(getattr(self.matcher, "selected_candidates", []))
        total_rows = len(self.matcher.bank) + len(self.matcher.journal)
        row_level = [
            item
            for item in candidates
            if len(item.bank_idxs) == 1
            and len(item.journal_idxs) == 1
            and item.metrics.total_diff_li == 0
            and item.processing_status.value == "自动确认"
        ]
        group_level = [
            item
            for item in candidates
            if item.metrics.total_diff_li == 0 and item not in row_level
            and item.processing_status.value in {"自动确认", "整组勾稽一致"}
        ]

        def side_indexes(items: list[Any], side: str) -> set[int]:
            return {index for item in items for index in getattr(item, f"{side}_idxs")}

        def covered(items: list[Any]) -> int:
            return len(side_indexes(items, "bank")) + len(side_indexes(items, "journal"))

        confirmed = row_level + group_level
        coverage_rows = []
        for label, side, frame in (
            ("银行", "bank", self.matcher.bank),
            ("序时账", "journal", self.matcher.journal),
        ):
            indexes = side_indexes(confirmed, side)
            total_amount = sum((abs(int(value)) for value in frame["amount_decimal"]), 0)
            confirmed_amount = sum(abs(int(frame.loc[index, "amount_decimal"])) for index in indexes)
            coverage_rows.extend([
                (f"{label}已核对笔数覆盖率", len(indexes) / len(frame) if len(frame) else 0.0),
                (f"{label}已核对金额覆盖率", confirmed_amount / total_amount if total_amount else 0.0),
            ])

        status_counts: Dict[str, int] = {}
        risk_counts: Dict[str, int] = {}
        status_amounts: Dict[str, Decimal] = {}
        risk_amounts: Dict[str, Decimal] = {}
        for item in candidates:
            status = item.processing_status.value
            risk = item.risk_level.value
            amount = abs(
                PrecisionEngine.from_integer_li(item.metrics.group_amount_li)
            )
            status_counts[status] = status_counts.get(status, 0) + 1
            risk_counts[risk] = risk_counts.get(risk, 0) + 1
            status_amounts[status] = status_amounts.get(status, Decimal("0")) + amount
            risk_amounts[risk] = risk_amounts.get(risk, Decimal("0")) + amount

        bank_unmatched = self.matcher.bank[~self.matcher.bank["matched"]]
        journal_unmatched = self.matcher.journal[~self.matcher.journal["matched"]]
        bank_unmatched_amount = sum(
            (abs(Decimal(str(value))) for value in bank_unmatched["amount"]),
            Decimal("0"),
        )
        journal_unmatched_amount = sum(
            (abs(Decimal(str(value))) for value in journal_unmatched["amount"]),
            Decimal("0"),
        )

        overall_control = getattr(self.matcher, "overall_control", None)
        overall_limited = bool(
            overall_control is not None and overall_control.scope_limited
        )
        scope_item_names = {
            "日期范围",
            "核对账户",
            "核对币种",
            "金额口径",
            "金额方向",
            "金额合计",
            "总体余额控制",
            "数据入口",
        }
        range_items = (
            [
                item
                for item in self.precheck_report.items
                if item.status == "疑点" and item.name in scope_item_names
            ]
            if self.precheck_report is not None and overall_limited
            else []
        )
        range_status = (
            "范围受限"
            if overall_limited
            else "范围可用"
            if self.precheck_report is not None
            else "范围未验证"
        )
        balance_integrity_limited = bool(
            getattr(self.matcher, "balance_integrity_limited", False)
        )
        affected_windows = (
            tuple(overall_control.affected_windows)
            if overall_control is not None
            else ()
        )
        if balance_integrity_limited and not overall_limited and affected_windows:
            break_days = "、".join(
                f"{window[1]:%Y-%m-%d}" for window in affected_windows
            )
            system_conclusion = (
                f"已完成自动分析；余额在{break_days}断档，"
                "仅断档窗口内关系降为疑点，其余关系按证据正常确认"
            )
        elif balance_integrity_limited:
            system_conclusion = "已完成自动分析；总体余额控制异常，相关关系已降为疑点"
        elif overall_limited:
            system_conclusion = "已完成自动分析；总体资料尚未闭合，具体关系已按范围限制分流"
        else:
            system_conclusion = "核对已自动完成；疑点已分级列示，不依赖人工填写"
        high_or_unknown = (
            risk_counts.get("高风险", 0) + risk_counts.get("范围未知", 0)
        )
        action1 = (
            "优先查看高风险及范围未知事项的业务依据"
            if high_or_unknown
            else "未发现高风险或范围未知事项，保留自动核对成果"
        )
        action2 = (
            "核查银行侧和日记账侧待查记录的截止期及入账依据"
            if len(bank_unmatched) or len(journal_unmatched)
            else "两侧记录均已形成自动关系或组级结论"
        )
        action3 = (
            "结合余额差异和连续性异常定位未入账或重复入账"
            if balance_check_possible
            else "本次无双方可用余额，余额核对未实施"
        )
        run_parameters = getattr(self.matcher, "run_parameters", {})
        combination_search = run_parameters.get("combination_search", {})
        candidate_search = run_parameters.get("candidate_search", {})
        stage_labels = {
            "business_group": "业务完整组",
            "whitelist": "白名单",
            "exact": "精确",
            "tolerance": "容差",
            "atomic_voucher": "完整凭证",
            "generic_combination": "通用组合",
        }
        stage_truncations = {
            stage: int(stats.get("truncated_source_rows", 0) or 0)
            for stage, stats in candidate_search.items()
            if isinstance(stats, dict)
        }
        generic_sources = int(
            combination_search.get("generic_source_rows", 0) or 0
        )
        fully_searched = int(
            combination_search.get("fully_searched_source_rows", 0) or 0
        )
        generic_truncated = int(
            combination_search.get("truncated_source_rows", 0) or 0
        )
        budget_exhausted = int(
            combination_search.get("budget_exhausted_source_rows", 0) or 0
        )
        depth_limited = int(
            combination_search.get("depth_limited_source_rows", 0) or 0
        )
        node_exhausted = int(
            combination_search.get("node_budget_exhausted_source_rows", 0) or 0
        )
        task_timeouts = int(
            combination_search.get("task_timeout_source_rows", 0) or 0
        )
        global_timeouts = int(
            combination_search.get(
                "global_timeout_unprocessed_source_rows",
                0,
            ) or 0
        )
        worker_failures = int(
            combination_search.get("worker_failure_source_rows", 0) or 0
        )
        has_ambiguity = any(
            bool(getattr(candidate, "is_ambiguous", False))
            for candidate in getattr(self.matcher, "candidates", ())
        )
        range_limited = bool(
            budget_exhausted
            or generic_truncated
            or any(stage_truncations.values())
        )
        if range_limited:
            search_completeness = "范围受限"
        elif has_ambiguity:
            search_completeness = "存在多解"
        else:
            search_completeness = "预算内完成"
        search_coverage = (
            f"通用来源{generic_sources}笔；"
            f"预算内完整搜索{fully_searched}笔；"
            f"预算受限{budget_exhausted}笔"
        )
        limitation_reasons = []
        for stage, label in stage_labels.items():
            if stage == "generic_combination":
                continue
            count = stage_truncations.get(stage, 0)
            if count:
                limitation_reasons.append(f"{label}截断{count}笔")
        if generic_truncated:
            limitation_reasons.append(
                f"通用组合候选截断{generic_truncated}笔"
            )
        if depth_limited:
            limitation_reasons.append(f"深度限制{depth_limited}笔")
        if node_exhausted:
            limitation_reasons.append(f"节点预算{node_exhausted}笔")
        if task_timeouts:
            limitation_reasons.append(f"单任务超时{task_timeouts}笔")
        if global_timeouts:
            limitation_reasons.append(f"全局超时未处理{global_timeouts}笔")
        if worker_failures:
            limitation_reasons.append(f"工作进程故障{worker_failures}笔")
        if has_ambiguity:
            limitation_reasons.append("存在多个可行对应关系")
        search_reason = "；".join(limitation_reasons) or "未触发搜索限制"
        search_explanation = (
            f"候选搜索未穷尽；{search_reason}。"
            "未找到对应不代表不存在组合。"
            if range_limited or has_ambiguity
            else "通用来源均在所列预算内完成。"
        )
        rows = [
            ("系统结论", system_conclusion),
            ("核对范围", range_status),
            ("范围说明", "；".join(item.name for item in range_items) or "未发现范围疑点"),
            ("银行有效交易笔数", len(self.matcher.bank)),
            ("日记账有效交易笔数", len(self.matcher.journal)),
            ("逐笔精确匹配率", covered(row_level) / total_rows if total_rows else 0.0),
            ("组级勾稽率", covered(group_level) / total_rows if total_rows else 0.0),
            ("自动完成率", 1.0 if total_rows else 0.0),
            ("自动完成率说明", "表示程序完成分析，包含疑点归集，不等于核对一致比例。"),
            ("匹配率口径", "逐笔及组级比例以两侧有效行数之和为分母；分侧覆盖率以各侧有效行数为分母。仅计入零差额的自动确认或整组勾稽一致，疑点和自动归集不计入。"),
            ("金额覆盖率口径", "各侧已核对记录金额绝对值之和÷该侧全部有效记录金额绝对值之和，收支不抵销；分母为零时记为0。"),
            ("组合搜索完整性", search_completeness),
            ("组合搜索覆盖", search_coverage),
            ("组合搜索受限原因", search_reason),
            ("组合搜索说明", search_explanation),
            *coverage_rows,
            ("自动确认组数", status_counts.get("自动确认", 0)),
            ("自动确认金额", float(status_amounts.get("自动确认", Decimal("0")))),
            ("整组勾稽组数", status_counts.get("整组勾稽一致", 0)),
            ("整组勾稽金额", float(status_amounts.get("整组勾稽一致", Decimal("0")))),
            ("自动归集事项数", status_counts.get("自动归集事项", 0)),
            ("自动归集金额", float(status_amounts.get("自动归集事项", Decimal("0")))),
            ("疑点事项数", status_counts.get("疑点事项", 0)),
            ("疑点事项金额", float(status_amounts.get("疑点事项", Decimal("0")))),
            ("低风险事项数", risk_counts.get("低风险", 0)),
            ("低风险事项金额", float(risk_amounts.get("低风险", Decimal("0")))),
            ("中风险事项数", risk_counts.get("中风险", 0)),
            ("中风险事项金额", float(risk_amounts.get("中风险", Decimal("0")))),
            ("高风险事项数", risk_counts.get("高风险", 0)),
            ("高风险事项金额", float(risk_amounts.get("高风险", Decimal("0")))),
            ("范围未知事项数", risk_counts.get("范围未知", 0)),
            ("范围未知事项金额", float(risk_amounts.get("范围未知", Decimal("0")))),
            *(
                [
                    ("中风险抽样总数", _sampling_stats["medium_total"]),
                    ("中风险抽中待核查数", _sampling_stats["sampled"]),
                ]
                if (
                    _sampling_stats := getattr(
                        self.matcher, "medium_sampling_stats", None
                    )
                )
                else []
            ),
            ("银行侧待查笔数", len(bank_unmatched)),
            ("银行侧待查金额", float(bank_unmatched_amount)),
            ("日记账侧待查笔数", len(journal_unmatched)),
            ("日记账侧待查金额", float(journal_unmatched_amount)),
            ("余额核对", "已实施" if balance_check_possible else "余额核对未实施"),
            ("待处理事项数", 0),
            ("已关注事项数", 0),
            ("已处理事项数", 0),
            ("无需处理事项数", 0),
            ("实际执行重要性水平", float(config.performance_materiality)),
            ("明显微小错报临界值", float(config.clearly_trivial_threshold)),
            ("建议动作1", action1),
            ("建议动作2", action2),
            ("建议动作3", action3),
        ]
        warning = self.initial_balance_warning
        if balance_check_possible and warning is not None:
            rows[2:2] = [
                ("银行期初余额", float(warning.bank_initial)),
                ("日记账期初余额", float(warning.journal_initial)),
                ("期初余额差额", float(warning.diff)),
                ("期初余额状态", "不一致" if warning.has_warning else "一致"),
            ]
        if not candidates and any(
            frame["match_id"].fillna("").ne("").any()
            for frame in (self.matcher.bank, self.matcher.journal)
        ):
            unavailable = "旧结果未保留核对状态依据，无法计算"
            rows = [(name, unavailable if name.endswith("覆盖率") or name in {"逐笔精确匹配率", "组级勾稽率"} else value) for name, value in rows]
        return pd.DataFrame(rows, columns=["项目", "数值"])

    def _build_issue_table(self) -> pd.DataFrame:
        columns = [
            "系统结论", "风险等级", "判断依据", "批次核查线索", "建议动作", "事项类型",
            "月份", "匹配类型", "银行笔数", "日记账笔数", "组金额", "差异金额",
            "银行原文件行号", "日记账原文件行号", "银行日期", "日记账日期",
            "银行金额", "日记账金额", "银行摘要", "日记账摘要",
            "后续状态", "处理说明", "调整凭证号", "责任人", "处理日期",
            "匹配ID", "候选ID", "综合可信度", "金额分", "日期分", "文字分",
            "结构分", "差异池ID",
        ]
        rows: list[dict[str, Any]] = []
        for candidate in getattr(self.matcher, "selected_candidates", []):
            risk = candidate.risk_level.value
            if (
                candidate.processing_status.value not in {"疑点事项", "自动归集事项"}
                and risk not in {"高风险", "范围未知"}
            ):
                continue
            sampling_mark = candidate.evidence.get("medium_sampling", "")
            action = self._suggested_action(risk)
            reason_text = candidate.processing_reason
            if risk == "中风险" and sampling_mark == "抽中待核查":
                reason_text += "；中风险等距抽样：抽中（样本量与高风险笔数一致）"
            elif risk == "中风险" and sampling_mark == "未抽中留存备查":
                action = "等距抽样未抽中，随同低风险留存备查，本次不留人工"
                reason_text += "；中风险等距抽样：未抽中"
            row = {
                "系统结论": candidate.processing_status.value,
                "风险等级": risk,
                "判断依据": reason_text,
                "批次核查线索": candidate.evidence.get("batch_review_hint", ""),
                "建议动作": action,
                "事项类型": (
                    "批次差异"
                    if candidate.evidence.get("batch_difference")
                    else "匹配或差异事项"
                ),
                "月份": min(candidate.bank_dates + candidate.journal_dates).strftime("%Y-%m")
                if candidate.bank_dates or candidate.journal_dates else "",
                "匹配类型": self._match_type_name(candidate.match_type),
                "银行笔数": len(candidate.bank_idxs),
                "日记账笔数": len(candidate.journal_idxs),
                "组金额": float(PrecisionEngine.from_integer_li(candidate.metrics.group_amount_li)),
                "差异金额": float(PrecisionEngine.from_integer_li(candidate.metrics.total_diff_li)),
                "后续状态": "",
                "处理说明": "",
                "调整凭证号": "",
                "责任人": "",
                "处理日期": "",
                "匹配ID": candidate.final_match_id,
                "候选ID": candidate.candidate_id,
                "综合可信度": candidate.scores.total,
                "金额分": candidate.scores.amount,
                "日期分": candidate.scores.date,
                "文字分": candidate.scores.text,
                "结构分": candidate.scores.structure,
                "差异池ID": "；".join(candidate.evidence.get("difference_pool_ids", [])),
            }
            row.update(self._candidate_snapshots(candidate))
            rows.append(row)

        existing_pool_ids = {row.get("差异池ID") for row in rows}
        for pool in getattr(self.matcher, "difference_pools", []):
            if pool.risk_level.value not in {"高风险", "范围未知"}:
                continue
            if pool.pool_id in existing_pool_ids:
                continue
            rows.append(
                {
                    "系统结论": pool.processing_status.value,
                    "风险等级": pool.risk_level.value,
                    "判断依据": pool.processing_reason,
                    "建议动作": self._suggested_action(pool.risk_level.value),
                    "事项类型": "月度差异池",
                    "月份": pool.month,
                    "匹配类型": pool.pool_type.value,
                    "组金额": float(PrecisionEngine.from_integer_li(pool.total_diff_li)),
                    "差异金额": float(PrecisionEngine.from_integer_li(pool.total_diff_li)),
                    "后续状态": "",
                    "处理说明": "",
                    "调整凭证号": "",
                    "责任人": "",
                    "处理日期": "",
                    "差异池ID": pool.pool_id,
                }
            )
        return pd.DataFrame(rows, columns=columns)

    def _build_business_event_table(self) -> pd.DataFrame:
        rows = []
        for event in getattr(self.matcher, "business_events", []):
            frame = self.matcher.bank if event.source == "bank" else self.matcher.journal
            source_name = "银行流水" if event.source == "bank" else "银行存款序时账"
            for sequence, index in enumerate(event.row_idxs, 1):
                row = frame.loc[index]
                rows.append({
                    "业务链ID": event.event_id,
                    "业务类型": event.event_type,
                    "来源": source_name,
                    "链内顺序": sequence,
                    "原文件行号": int(row.get("original_file_row", row.get("original_idx", index))),
                    "日期": row.get("date", ""),
                    "金额": float(row.get("amount", 0)),
                    "摘要": row.get("summary", ""),
                    "关系公式": event.relationship_formula,
                    "最终净影响": float(PrecisionEngine.from_integer_li(event.net_amount_li)),
                    "业务依据": event.evidence_basis,
                    "是否跨期": "是" if event.is_cross_period else "否",
                    "系统结论": event.review_status,
                })
        return pd.DataFrame(rows, columns=["业务链ID", "业务类型", "来源", "链内顺序", "原文件行号", "日期", "金额", "摘要", "关系公式", "最终净影响", "业务依据", "是否跨期", "系统结论"])

    def _build_cutoff_table(self, events: pd.DataFrame) -> pd.DataFrame:
        """保留跨期业务链，并独立披露所有跨月对应关系，包括小额自动确认。"""
        rows = []
        for candidate in getattr(self.matcher, "selected_candidates", []):
            dates = candidate.bank_dates + candidate.journal_dates
            if len({(day.year, day.month) for day in dates}) <= 1:
                continue
            row = {
                "事项类型": "跨期对应关系",
                "匹配ID": candidate.final_match_id,
                "候选ID": candidate.candidate_id,
                "业务类型": self._match_type_name(candidate.match_type),
                "是否跨期": "是",
                "系统结论": candidate.processing_status.value,
                "业务依据": candidate.processing_reason,
                "建议动作": "核查两侧入账期间及截止性；金额对应不代表期间正确",
            }
            row.update(self._candidate_snapshots(candidate))
            rows.append(row)
        cross_events = events.loc[events["是否跨期"] == "是"].copy()
        cross_events["事项类型"] = "跨期业务链"
        if not rows:
            return cross_events.reset_index(drop=True)
        return pd.concat([pd.DataFrame(rows), cross_events], ignore_index=True).fillna("")

    def _build_business_clue_table(self) -> pd.DataFrame:
        rows = []
        for clue in getattr(self.matcher, "business_clues", []):
            frame = self.matcher.bank if clue.source == "bank" else self.matcher.journal
            source_name = "银行流水" if clue.source == "bank" else "银行存款序时账"
            for index in clue.row_idxs:
                row = frame.loc[index]
                rows.append({
                    "线索ID": clue.clue_id,
                    "线索类型": clue.clue_type,
                    "来源": source_name,
                    "原文件行号": int(row.get("original_file_row", row.get("original_idx", index))),
                    "日期": row.get("date", ""),
                    "金额": float(row.get("amount", 0)),
                    "摘要": row.get("summary", ""),
                    "判断依据": clue.reason,
                    "处理说明": "保留全部原始记录，结合回单、凭证和对方资料核查",
                })
        return pd.DataFrame(rows, columns=["线索ID", "线索类型", "来源", "原文件行号", "日期", "金额", "摘要", "判断依据", "处理说明"])

    def _business_chain_rows(self, source: str) -> dict[int, list[str]]:
        """原文件行号 → 所属同侧业务链事件号，用于待查表交叉引用。"""
        frame = self.matcher.bank if source == "bank" else self.matcher.journal
        lookup: dict[int, list[str]] = {}
        for event in getattr(self.matcher, "business_events", []):
            if event.source != source:
                continue
            for index in event.row_idxs:
                row = frame.loc[index]
                file_row = int(row.get("original_file_row", row.get("original_idx", index)))
                lookup.setdefault(file_row, []).append(event.event_id)
        return lookup

    def _decorate_unmatched(self, table: pd.DataFrame, side: str) -> pd.DataFrame:
        result = table.copy()
        result.insert(0, "建议动作", "结合对侧记录和截止期继续核查")
        result.insert(0, "判断依据", "程序未找到足以建立关系的候选")
        result.insert(0, "风险等级", "范围未知")
        result.insert(0, "系统结论", f"{side}未找到候选")
        chain_rows = self._business_chain_rows("bank" if side == "银行侧" else "journal")
        if chain_rows and "原文件行号" in result.columns:
            linked = result["原文件行号"].map(
                lambda value: chain_rows.get(int(value)) if pd.notna(value) else None
            )
            mask = linked.notna()
            result.loc[mask, "判断依据"] = linked[mask].map(
                lambda ids: f"已纳入同侧业务链{'、'.join(ids)}，组成和净影响见退款冲销重付表"
            )
            result.loc[mask, "建议动作"] = "按业务链整体核查退回、冲销或重付凭证，不单行核对"
        return result

    def _apply_report_presentation(self, workbook: Any) -> None:
        technical_sheets = {"运行参数", "大模型辅助明细", "解析异常明细"}
        technical_columns = {
            "候选ID", "阶段", "最终状态", "处理原因", "综合可信度",
            "金额分", "日期分", "文字分", "结构分", "银_日期", "账_日期",
            "文字支持", "文字冲突", "大模型判断", "差异池ID", "是否使用大模型",
            "纳入风险池", "匹配候选ID", "银行收入", "银行支出", "银行净额",
            "日记账收入", "日记账支出", "日记账净额", "组金额", "收入差额",
            "支出差额",
        }
        issue_hidden_columns = {
            "银行原文件行号", "日记账原文件行号", "银行日期", "日记账日期",
            "银行金额", "日记账金额", "调整凭证号", "责任人", "处理日期", "匹配ID",
        }
        amount_keywords = ("金额", "合计", "收入", "支出", "净额", "差额", "余额", "组金额")
        percent_headers = {
            "逐笔精确匹配率", "组级勾稽率", "自动完成率",
            "银行已核对笔数覆盖率", "银行已核对金额覆盖率",
            "序时账已核对笔数覆盖率", "序时账已核对金额覆盖率",
        }
        for sheet in workbook.worksheets:
            if sheet.title in technical_sheets:
                sheet.sheet_state = "hidden"
            business_columns = [cell.column for cell in sheet[1] if cell.value is not None]
            first_business_column = min(business_columns) if business_columns else 1
            if first_business_column > 1:
                for column in range(1, first_business_column):
                    sheet.column_dimensions[get_column_letter(column)].hidden = True
            sheet.freeze_panes = f"{get_column_letter(first_business_column)}2"
            if sheet.max_row >= 1 and sheet.max_column >= 1:
                sheet.auto_filter.ref = (
                    f"{get_column_letter(first_business_column)}1:"
                    f"{get_column_letter(sheet.max_column)}{sheet.max_row}"
                )
            for cell in sheet[1]:
                if cell.value in technical_columns:
                    sheet.column_dimensions[cell.column_letter].hidden = True
                if sheet.title == "疑点事项" and cell.value in issue_hidden_columns:
                    sheet.column_dimensions[cell.column_letter].hidden = True
                width = 14
                if cell.value in {"系统结论", "判断依据", "业务分组依据", "其他可能对应", "建议动作", "摘要", "银行摘要", "日记账摘要", "说明"}:
                    width = 28
                sheet.column_dimensions[cell.column_letter].width = width
                header = str(cell.value or "")
                # 空表只有表头：不触碰第2行，避免凭空造出一行全空单元格。
                if sheet.max_row < 2:
                    continue
                for data_cell in sheet.iter_cols(
                    min_col=cell.column,
                    max_col=cell.column,
                    min_row=2,
                    max_row=sheet.max_row,
                ):
                    for item in data_cell:
                        item.alignment = Alignment(vertical="top", wrap_text=True)
                        if any(keyword in header for keyword in amount_keywords):
                            item.number_format = '#,##0.00'
                        if isinstance(item.value, str) and not sheet.column_dimensions[cell.column_letter].hidden:
                            # 中文通常占两字符宽，按可见列估算换行，避免依据文字被固定行高遮住。
                            lines = sum(max(1, (sum(2 if ord(c) > 127 else 1 for c in line) + width - 1) // width) for line in item.value.split("\n"))
                            sheet.row_dimensions[item.row].height = min(409, max(sheet.row_dimensions[item.row].height or 20, lines * 16 + 4))
            if sheet.title == "核对结论":
                headers = {cell.value: cell.column for cell in sheet[1] if cell.value is not None}
                item_column = headers.get("项目", first_business_column)
                value_column = headers.get("数值", item_column + 1)
                sheet.column_dimensions[
                    get_column_letter(item_column)
                ].width = 24
                sheet.column_dimensions[
                    get_column_letter(value_column)
                ].width = 58
                for row in range(2, sheet.max_row + 1):
                    if sheet.cell(row=row, column=item_column).value in percent_headers:
                        sheet.cell(row=row, column=value_column).number_format = "0.00%"
                    value = sheet.cell(row=row, column=value_column).value
                    text_width = sum(2 if ord(c) > 127 else 1 for c in str(value or ""))
                    sheet.row_dimensions[row].height = min(409, max(32 if len(str(value or "")) > 20 else 20, ((text_width + 51) // 52) * 16 + 4))
            if sheet.title == "输入检查":
                headers = {
                    cell.value: cell.column
                    for cell in sheet[1]
                    if cell.value is not None
                }
                widths = {
                    "检查项目": 20,
                    "银行流水结果": 20,
                    "银行日记账结果": 20,
                    "双方比较结果": 24,
                    "状态": 12,
                    "说明": 52,
                }
                for header, width in widths.items():
                    column = headers.get(header)
                    if column is not None:
                        sheet.column_dimensions[
                            get_column_letter(column)
                        ].width = width
                for row in range(2, sheet.max_row + 1):
                    longest_text = max(
                        (
                            len(str(sheet.cell(row=row, column=column).value or ""))
                            for column in headers.values()
                        ),
                        default=0,
                    )
                    sheet.row_dimensions[row].height = (
                        36 if longest_text > 18 else 22
                    )

        if "疑点事项" in workbook.sheetnames:
            sheet = workbook["疑点事项"]
            headers = {cell.value: cell.column for cell in sheet[1]}
            status_column = headers.get("后续状态")
            if status_column is not None:
                validation = DataValidation(
                    type="list",
                    formula1='"已关注,已处理,无需处理"',
                    allow_blank=True,
                )
                sheet.add_data_validation(validation)
                letter = get_column_letter(status_column)
                validation.add(f"{letter}2:{letter}{max(2, sheet.max_row)}")

            if "核对结论" in workbook.sheetnames:
                item_column = headers.get("系统结论")
                if status_column is not None and item_column is not None:
                    last_row = max(2, sheet.max_row)
                    status_letter = get_column_letter(status_column)
                    item_letter = get_column_letter(item_column)
                    status_range = (
                        f"'疑点事项'!${status_letter}$2:"
                        f"${status_letter}${last_row}"
                    )
                    item_range = (
                        f"'疑点事项'!${item_letter}$2:"
                        f"${item_letter}${last_row}"
                    )
                    formulas = {
                        "待处理事项数": (
                            f'=COUNTIFS({item_range},"<>",{status_range},"")'
                        ),
                        "已关注事项数": f'=COUNTIF({status_range},"已关注")',
                        "已处理事项数": f'=COUNTIF({status_range},"已处理")',
                        "无需处理事项数": f'=COUNTIF({status_range},"无需处理")',
                    }
                    summary = workbook["核对结论"]
                    summary_headers = {
                        cell.value: cell.column for cell in summary[1]
                    }
                    summary_item_column = summary_headers.get("项目", 1)
                    summary_value_column = summary_headers.get("数值", 2)
                    for row in range(2, summary.max_row + 1):
                        item = summary.cell(
                            row=row,
                            column=summary_item_column,
                        ).value
                        if item in formulas:
                            summary.cell(
                                row=row,
                                column=summary_value_column,
                            ).value = formulas[item]
                    workbook.calculation.calcMode = "auto"
                    workbook.calculation.fullCalcOnLoad = True
                    workbook.calculation.forceFullCalc = True

    def build_report_tables(
        self,
        config: MatcherConfig,
        date_format: str = "auto",
    ) -> dict[str, pd.DataFrame]:
        """在写入 Excel 前构造所有可单独检查的结构化表。"""
        balance_possible, _ = self._prepare_initial_balance()
        bank_has_balance = self._has_balance_data(
            self.matcher.bank,
            self.bank_mapping,
        )
        journal_source = (
            self.raw_journal
            if self.raw_journal is not None
            else self.matcher.journal
        )
        journal_has_balance = self._has_balance_data(
            journal_source,
            self.journal_mapping,
        )
        daily, monthly = self._build_daily_and_monthly_tables()
        groups = self._build_match_group_table()
        events = self._build_business_event_table()
        row_mask = (
            (groups["银行笔数"] == 1) & (groups["日记账笔数"] == 1)
            if not groups.empty
            else pd.Series(dtype=bool)
        )
        tables = {
            "核对结论": self._build_business_summary_table(
                config,
                balance_check_possible=balance_possible,
            ),
            "疑点事项": self._build_issue_table(),
            "自动归集事项": self._build_trivial_table(),
            "银行侧待查": self._decorate_unmatched(
                self._build_unmatched_table("bank"), "银行侧"
            ),
            "日记账侧待查": self._decorate_unmatched(
                self._build_unmatched_table("journal"), "日记账侧"
            ),
            "逐笔匹配": groups.loc[row_mask].reset_index(drop=True)
            if not groups.empty else groups.copy(),
            "整组勾稽": groups.loc[~row_mask].reset_index(drop=True)
            if not groups.empty else groups.copy(),
            "匹配组成": self._build_match_component_table(),
            "退款冲销重付": events,
            "手续费及净额": groups.loc[groups["类型"] == "手续费净额"].reset_index(drop=True) if not groups.empty else groups.copy(),
            "截止性差异": self._build_cutoff_table(events),
            "重复线索": self._build_business_clue_table(),
        }
        alternatives = self._build_alternative_component_table()
        if not alternatives.empty:
            tables["其他可能对应明细"] = alternatives
        if self.precheck_report is not None:
            tables["输入检查"] = self.precheck_report.to_dataframe()
            tables["运行资料与映射"] = self.precheck_report.source_dataframe()
            tables["数据入口处置"] = self.precheck_report.population_dataframe()
        tables["月度统计"] = monthly
        tables["每日统计"] = daily
        tables.update(
            self._build_balance_tables(
                bank_has_balance=bank_has_balance,
                journal_has_balance=journal_has_balance,
            )
        )
        llm_table = self._build_llm_table()
        if not llm_table.empty:
            tables["大模型辅助明细"] = llm_table
        if self.error_collector and self.error_collector.has_errors():
            errors = self.error_collector.get_all_errors()
            if errors:
                error_table = pd.DataFrame(errors).rename(
                    columns={
                        "type": "异常类型",
                        "source_type": "来源",
                        "row": "原文件行号",
                        "column": "字段",
                        "original_value": "原值",
                        "error": "原因",
                    }
                )
                tables["解析异常明细"] = error_table
        tables["运行参数"] = self._build_parameter_table(
            config,
            date_format,
        )
        return {
            name: self._safe_table(frame)
            for name, frame in tables.items()
        }

    def generate_report(
        self,
        output_path: str,
        config: Optional[MatcherConfig] = None,
        bank_path: Optional[str] = None,
        journal_path: Optional[str] = None,
        date_format: str = "auto",
    ) -> None:
        """生成组级清晰、可复核且防公式注入的 Excel 报告。"""
        del bank_path, journal_path
        report_started = time.perf_counter()
        effective_config = config or self.matcher.config
        table_started = time.perf_counter()
        self._log("开始构造报告表格")
        tables = self.build_report_tables(
            effective_config,
            date_format=date_format,
        )
        total_rows = sum(len(table) for table in tables.values())
        self._log(
            f"报告表格构造完成：{len(tables):,} 个工作表，"
            f"共 {total_rows:,} 行，耗时 {time.perf_counter() - table_started:.1f} 秒"
        )
        write_started = time.perf_counter()
        self._log(
            f"开始写入 Excel 工作簿：{len(tables):,} 个工作表，"
            f"共 {total_rows:,} 行"
        )
        make_excel(
            list(tables.items()),
            output_path,
            theme="deep-navy",
        )
        self._log(
            f"工作簿初次写入完成，耗时 {time.perf_counter() - write_started:.1f} 秒"
        )
        presentation_started = time.perf_counter()
        self._log("开始应用审计报告排版和复核标记")
        workbook = load_workbook(output_path)
        has_warning = bool(
            self.initial_balance_warning
            and self.initial_balance_warning.has_warning
        )
        if "核对结论" in workbook.sheetnames:
            self._postprocess_summary(
                workbook["核对结论"],
                has_warning,
            )
        for detail_name in ("逐笔匹配", "整组勾稽", "疑点事项", "自动归集事项"):
            if detail_name in workbook.sheetnames:
                self._postprocess_details(workbook[detail_name])
        for sheet_name in (
            "每日统计",
            "月度统计",
            "余额差异明细",
            "余额连续性异常",
        ):
            if sheet_name in workbook.sheetnames:
                self._postprocess_diff_columns(workbook[sheet_name])

        self._apply_report_presentation(workbook)
        workbook.save(output_path)
        self._log(
            f"报告排版保存完成：排版耗时 "
            f"{time.perf_counter() - presentation_started:.1f} 秒，"
            f"报告总耗时 {time.perf_counter() - report_started:.1f} 秒"
        )
