import gc
import re
import threading
from pathlib import Path
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import Optional, List, Dict, Any, Callable, Union, Tuple

import pandas as pd
import numpy as np

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

from precision_engine import PrecisionEngine
from utils import (
    parse_date,
    clean_amount,
    normalize_amount_text,
    clean_excel_string,
)


_BANK_DIRECTION_SIGNS = {
    '贷': 1, '贷方': 1, '收': 1, '收入': 1, '来账': 1, '收款': 1,
    '入': 1, '转入': 1, '存入': 1, '进': 1,
    'CREDIT': 1, 'C': 1, 'CR': 1, 'INCOME': 1, 'IN': 1,
    '借': -1, '借方': -1, '支': -1, '支出': -1, '往账': -1, '付': -1,
    '出': -1, '转出': -1, '取出': -1, '退': -1, '付款': -1,
    'DEBIT': -1, 'D': -1, 'DR': -1, 'EXPENSE': -1, 'OUT': -1,
}

_JOURNAL_DIRECTION_SIGNS = {
    '借': 1, '借方': 1, '收': 1, '收入': 1, '来账': 1, '收款': 1,
    '入': 1, '转入': 1, '存入': 1, '进': 1,
    'DEBIT': 1, 'D': 1, 'DR': 1, 'INCOME': 1, 'IN': 1,
    '贷': -1, '贷方': -1, '支': -1, '支出': -1, '往账': -1, '付': -1,
    '出': -1, '转出': -1, '取出': -1, '退': -1, '付款': -1,
    'CREDIT': -1, 'C': -1, 'CR': -1, 'EXPENSE': -1, 'OUT': -1,
}


def direction_sign(value: Any, source_type: str) -> Optional[int]:
    """按银行或企业日记账口径解释方向。"""
    if source_type == 'bank':
        signs = _BANK_DIRECTION_SIGNS
    elif source_type == 'journal':
        signs = _JOURNAL_DIRECTION_SIGNS
    else:
        raise ValueError(f"未知数据来源: {source_type}")
    return signs.get(str(value).strip().upper())


def parse_source_amount(
    value: Any,
    source_type: str,
    allow_suffix_sign: bool = True,
) -> Optional[Decimal]:
    """解析金额，并让 CR/DR/借/贷后缀遵循数据来源口径。"""
    if not allow_suffix_sign or not isinstance(value, str):
        return clean_amount(value, allow_suffix_sign=False)

    text = value.strip()
    suffix_match = re.search(r'(CR|DR|借|贷)\s*$', text, flags=re.IGNORECASE)
    if not suffix_match:
        return clean_amount(value, allow_suffix_sign=False)

    amount = clean_amount(text[:suffix_match.start()].strip(), allow_suffix_sign=False)
    sign = direction_sign(suffix_match.group(1), source_type)
    if amount is None or sign is None:
        return None
    return (amount * Decimal(sign)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


# ==========================================
# 1. 数据清洗器 (DataCleaner)
# ==========================================

class DataCleaner:
    """数据清洗器 - 自动处理日记账中的无效行
    
    自动检测并跳过：
    1. 开头的合并单元格标题行
    2. 数据中间的汇总行（月合计、本期累计等）
    """
    
    SUMMARY_KEYWORDS: Tuple[str, ...] = (
        '合计', '累计', '小计', '总计', '大计', 
        '本期', '本日', '本月', '本年', '本季',
        '日计', '月计', '年计', '季计',
        '结转', '期初', '期末'
    )
    
    DATE_COL_KEYWORDS: Tuple[str, ...] = ('日期', 'date', '记账时间', 'time', '交易日期')
    VOUCHER_COL_KEYWORDS: Tuple[str, ...] = ('凭证', 'voucher', '凭证号', '凭证编号')
    
    def __init__(self, logger: Optional[Callable[[str], None]] = None):
        self.logger = logger
    
    def _log(self, msg: str) -> None:
        if self.logger:
            self.logger(msg)
    
    def detect_skip_rows(self, df: pd.DataFrame, max_scan: int = 5) -> int:
        """检测开头的合并单元格标题行数量
        
        判断标准：某行几乎全是NaN/空值，且只有1-2个非空单元格
        合并单元格标题行的特征：第一列有值，其他列全是None
        """
        if df.empty:
            return 0
        
        skip_count = 0
        for idx in range(min(max_scan, len(df))):
            row = df.iloc[idx]
            non_null_count = row.notna().sum()
            total_count = len(row)
            
            if total_count == 0:
                skip_count += 1
                continue
            
            if non_null_count <= 2 and non_null_count / total_count < 0.3:
                skip_count += 1
            else:
                break
        
        if skip_count > 0:
            self._log(f"数据清洗：检测到{skip_count}行标题行，将跳过")
        
        return skip_count
    
    def is_summary_row(self, row: pd.Series, date_col: str, voucher_col: str) -> bool:
        """判断某行是否为汇总行
        
        判断条件（同时满足）：
        1. 日期列不是有效日期格式（包含汇总关键词）
        2. 凭证编号列为空或NaN
        """
        if date_col not in row.index or voucher_col not in row.index:
            return False
        
        date_val = row[date_col]
        voucher_val = row[voucher_col]
        
        if pd.notna(voucher_val) and str(voucher_val).strip() != '':
            return False
        
        if pd.isna(date_val):
            return False
        
        date_str = str(date_val).strip()
        if not date_str:
            return False
        
        for keyword in self.SUMMARY_KEYWORDS:
            if keyword in date_str:
                return True
        
        return False
    
    def _is_summary_row_vectorized(self, df: pd.DataFrame, date_col: str, voucher_col: str) -> pd.Series:
        """向量化判断汇总行
        
        判断条件（同时满足）：
        1. 凭证编号列为空或NaN
        2. 日期列包含汇总关键词
        """
        if date_col not in df.columns or voucher_col not in df.columns:
            return pd.Series([False] * len(df), index=df.index)
        
        voucher_is_empty = df[voucher_col].isna()
        voucher_str = df[voucher_col].astype(str).str.strip()
        voucher_is_empty = voucher_is_empty | (voucher_str == '') | (voucher_str.str.lower() == 'nan')
        
        date_not_na = df[date_col].notna()
        date_str = df[date_col].astype(str).str.strip()
        date_not_empty = date_str != ''
        
        pattern = '|'.join(re.escape(kw) for kw in self.SUMMARY_KEYWORDS)
        date_contains_keyword = date_str.str.contains(pattern, case=False, na=False, regex=True)
        
        return voucher_is_empty & date_not_na & date_not_empty & date_contains_keyword
    
    def _auto_detect_columns(self, df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
        """自动检测日期列和凭证列"""
        date_col = None
        voucher_col = None
        
        col_map = {str(col).lower().strip(): col for col in df.columns}
        
        for kw in self.DATE_COL_KEYWORDS:
            for col_lower, col_orig in col_map.items():
                if kw.lower() in col_lower:
                    date_col = col_orig
                    break
            if date_col:
                break
        
        for kw in self.VOUCHER_COL_KEYWORDS:
            for col_lower, col_orig in col_map.items():
                if kw.lower() in col_lower:
                    voucher_col = col_orig
                    break
            if voucher_col:
                break
        
        return date_col, voucher_col
    
    def clean_data(
        self, 
        df: pd.DataFrame, 
        date_col: Optional[str] = None, 
        voucher_col: Optional[str] = None
    ) -> pd.DataFrame:
        """执行完整的数据清洗流程
        
        :param df: 原始数据
        :param date_col: 日期列名（可选，自动检测）
        :param voucher_col: 凭证列名（可选，自动检测）
        :return: 清洗后的DataFrame
        """
        if df.empty:
            return df
        
        if date_col is None or voucher_col is None:
            auto_date, auto_voucher = self._auto_detect_columns(df)
            date_col = date_col or auto_date
            voucher_col = voucher_col or auto_voucher
        
        if date_col is None or voucher_col is None:
            self._log("数据清洗：无法自动检测日期列或凭证列，跳过清洗")
            return df
        
        skip_rows = self.detect_skip_rows(df)
        if skip_rows > 0:
            df = df.iloc[skip_rows:].reset_index(drop=True)
        
        if df.empty:
            return df
        
        summary_mask = self._is_summary_row_vectorized(df, date_col, voucher_col)
        
        summary_count = summary_mask.sum()
        if summary_count > 0:
            df = df[~summary_mask].reset_index(drop=True)
            self._log(f"数据清洗：过滤{summary_count}行汇总记录（月合计/本期累计等）")
        
        return df


# ==========================================
# 2. 解析错误收集器 (ParseErrorCollector)
# ==========================================

class ParseErrorCollector:
    """收集解析过程中的错误和异常"""
    
    def __init__(self):
        self._errors: List[Dict[str, Any]] = []
        self._counts: Dict[str, int] = {
            '金额解析失败': 0,
            '方向解析失败': 0,
            '日期解析失败': 0,
            '被丢弃的汇总行': 0,
            '空日期行': 0
        }
        self._lock = threading.Lock()
    
    def record_amount_error(self, row_idx: int, original_value: Any, source_type: str, column: str) -> None:
        with self._lock:
            self._counts['金额解析失败'] += 1
            if len(self._errors) < 1000:
                self._errors.append({
                    'type': '金额解析失败',
                    'row': row_idx,
                    'original_value': str(original_value),
                    'source_type': source_type,
                    'column': column
                })

    def record_direction_error(
        self,
        row_idx: int,
        original_value: Any,
        source_type: str,
        column: str,
    ) -> None:
        with self._lock:
            self._counts['方向解析失败'] += 1
            if len(self._errors) < 1000:
                self._errors.append({
                    'type': '方向解析失败',
                    'row': row_idx,
                    'original_value': str(original_value),
                    'source_type': source_type,
                    'column': column,
                })
    
    def record_date_error(self, row_idx: int, original_value: Any, source_type: str, column: str) -> None:
        with self._lock:
            self._counts['日期解析失败'] += 1
            if len(self._errors) < 1000:
                self._errors.append({
                    'type': '日期解析失败',
                    'row': row_idx,
                    'original_value': str(original_value),
                    'source_type': source_type,
                    'column': column
                })
    
    def record_dropped_summary_row(self, row_idx: int, reason: str, original_data: Dict[str, Any]) -> None:
        with self._lock:
            self._counts['被丢弃的汇总行'] += 1
            if len(self._errors) < 1000:
                self._errors.append({
                    'type': '被丢弃的汇总行',
                    'row': row_idx,
                    'reason': reason,
                    'original_data': original_data
                })
    
    def record_empty_date_row(self, row_idx: int, source_type: str) -> None:
        with self._lock:
            self._counts['空日期行'] += 1
            if len(self._errors) < 1000:
                self._errors.append({
                    'type': '空日期行',
                    'row': row_idx,
                    'source_type': source_type
                })
    
    def has_errors(self) -> bool:
        with self._lock:
            return len(self._errors) > 0
    
    def get_all_errors(self) -> List[Dict[str, Any]]:
        with self._lock:
            return list(self._errors)
    
    def get_summary(self) -> Dict[str, Any]:
        with self._lock:
            return {
                '总计': sum(self._counts.values()),
                **dict(self._counts)
            }


# ==========================================
# 3. 数据加载器 (DataLoader)
# ==========================================

class DataLoader:
    def __init__(self, logger: Optional[Callable[[str], None]] = None,
                 error_collector: Optional[ParseErrorCollector] = None):
        self.logger = logger
        self.error_collector = error_collector

    def load_file(self, file_path: str, skiprows: int = 0) -> pd.DataFrame:
        """
        读取 Excel 或 CSV 文件，返回 DataFrame
        """

        path_obj = Path(file_path)
        if not path_obj.exists():
            raise FileNotFoundError(f"文件未找到: {file_path}")

        ext = path_obj.suffix.lower()
        
        # 编码列表（按优先级）
        encodings = ['utf-8-sig', 'gbk', 'latin1']
        
        if ext in ['.xlsx', '.xls']:
            # Excel不支持分块读取，直接读取
            df = pd.read_excel(file_path, dtype=object, skiprows=skiprows)
        elif ext == '.csv':
            loaded = False
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding, dtype=object, skiprows=skiprows)
                    loaded = True
                    break
                except UnicodeDecodeError:
                    continue
            if not loaded:
                raise UnicodeDecodeError("无法解码CSV文件，尝试了多种编码", b'', 0, 1, 'all encodings failed')
        else:
            raise ValueError("不支持的文件格式")
            
        return df

    def find_header_row(self, file_path: str, max_scan_rows: int = 10) -> int:
        """
        自动检测表头行位置
        """
        path_obj = Path(file_path)
        if not path_obj.exists():
            return 0
        
        ext = path_obj.suffix.lower()
        header_keywords = ['日期', 'Date', 'date', '摘要', 'Summary', 'summary', '金额', 'Amount', 'amount', 
                          '借方', 'Debit', 'debit', '贷方', 'Credit', 'credit', '余额', 'Balance', 'balance']
        
        try:
            if ext in ['.xlsx', '.xls']:
                # 一次性读取前 N 行，避免逐行重开文件
                df_scan = pd.read_excel(file_path, dtype=object, header=None, nrows=max_scan_rows)
                for i in range(len(df_scan)):
                    row_values = df_scan.iloc[i].astype(str).values
                    keyword_count = sum(1 for val in row_values if any(kw in val for kw in header_keywords))
                    if keyword_count >= 2:
                        if self.logger:
                            self.logger(f"自动检测到表头在第 {i + 1} 行")
                        return i
            elif ext == '.csv':
                encodings = ['utf-8-sig', 'gbk', 'latin1']
                for encoding in encodings:
                    try:
                        df_scan = pd.read_csv(file_path, encoding=encoding, dtype=object, header=None, nrows=max_scan_rows)
                        for i in range(len(df_scan)):
                            row_values = df_scan.iloc[i].astype(str).values
                            keyword_count = sum(1 for val in row_values if any(kw in val for kw in header_keywords))
                            if keyword_count >= 2:
                                if self.logger:
                                    self.logger(f"自动检测到表头在第 {i + 1} 行")
                                return i
                        break
                    except UnicodeDecodeError:
                        continue
        except Exception as e:
            if self.logger:
                self.logger(f"自动检测表头行时出错: {str(e)}")
        
        return 0

    def _process_amount_column(self, data: pd.DataFrame, col_name: Optional[str], 
                                  source_type: str = "") -> Tuple[pd.Series, List[Dict[str, Any]]]:
        """
        处理金额列的辅助方法。
        优化：避免 float 中间态，直接从清洗后的字符串转 Decimal。
        """
        parse_failures: List[Dict[str, Any]] = []
        
        if col_name and col_name in data.columns:
            def safe_to_decimal(raw_value: Any) -> Optional[Decimal]:
                normalized = normalize_amount_text(raw_value)
                if normalized is None:
                    return None
                if normalized == "" or normalized.lower() in {"nan", "none"}:
                    return None
                try:
                    return Decimal(normalized).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                except (ValueError, TypeError, InvalidOperation):
                    return None

            result = data[col_name].apply(safe_to_decimal)

            original_not_empty = data[col_name].notna() & \
                                 (data[col_name].astype(str).str.strip() != '') & \
                                 (data[col_name].astype(str).str.strip().str.lower() != 'nan')
            
            failed_mask = result.isna() & original_not_empty
            
            if failed_mask.any():
                failed_values = data.loc[failed_mask, col_name]
                cleaned_values = failed_values.apply(lambda x: clean_amount(x, allow_suffix_sign=False))
                
                for idx, val in cleaned_values.items():
                    if val is None:
                        original_idx = data.loc[idx, 'original_file_row'] if 'original_file_row' in data.columns else idx + 1
                        parse_failures.append({
                            'row': original_idx,
                            'original_value': data.loc[idx, col_name],
                            'source_type': source_type,
                            'column': col_name
                        })
                        if self.error_collector:
                            self.error_collector.record_amount_error(
                                original_idx, 
                                data.loc[idx, col_name],
                                source_type, 
                                col_name
                            )
                    else:
                        result.loc[idx] = val
            
            # 仅对原本为空的值补 0，解析失败的数据保留为 None 以便后续剔除并报警。
            empty_mask = result.isna() & (~original_not_empty)
            result.loc[empty_mask] = Decimal('0.00')
            
            return result, parse_failures
            
        return pd.Series(Decimal('0.00'), index=data.index), parse_failures

    @staticmethod
    def _is_non_empty_cell(value: Any) -> bool:
        """判断单元格是否包含有效内容。"""
        if pd.isna(value):
            return False
        text = str(value).strip()
        return text != "" and text.lower() != "nan"

    def _build_date_ffill_mask(self, data: pd.DataFrame, mapping: Dict[str, Any]) -> pd.Series:
        """
        为常见序时账“日期合并单元格/续行空日期”场景构建条件性前填掩码。

        仅在以下条件同时满足时允许沿用上一行日期：
        1. 当前日期单元格为空；
        2. 行内存在其他交易字段内容，说明这是业务续行而不是纯空白行；
        3. 前面已经存在可沿用的日期。
        """
        date_col = mapping.get('date')
        date_str = data[date_col].astype(str).str.strip()
        blank_date_mask = data[date_col].isna() | (date_str == '') | (date_str.str.lower() == 'nan')

        candidate_cols = []
        for key in ('summary', 'voucher', 'debit', 'credit', 'amount', 'direction', 'balance'):
            col_name = mapping.get(key)
            if col_name and col_name in data.columns and col_name != date_col:
                candidate_cols.append(col_name)

        has_transaction_content = pd.Series(False, index=data.index)
        for col_name in candidate_cols:
            has_transaction_content = has_transaction_content | data[col_name].apply(self._is_non_empty_cell)

        mode = mapping.get('mode', 'debit_credit')
        if mode == 'debit_credit':
            # 兼容常见账套：续行分录常常只保留借贷金额，日期和凭证号都留空。
            voucher_col = mapping.get('voucher')
            debit_col = mapping.get('debit')
            credit_col = mapping.get('credit')
            if debit_col in data.columns or credit_col in data.columns:
                voucher_blank = (
                    pd.Series(True, index=data.index)
                    if not voucher_col or voucher_col not in data.columns
                    else ~data[voucher_col].apply(self._is_non_empty_cell)
                )
                debit_has_value = (
                    pd.Series(False, index=data.index)
                    if not debit_col or debit_col not in data.columns
                    else data[debit_col].apply(self._is_non_empty_cell)
                )
                credit_has_value = (
                    pd.Series(False, index=data.index)
                    if not credit_col or credit_col not in data.columns
                    else data[credit_col].apply(self._is_non_empty_cell)
                )
                has_transaction_content = has_transaction_content | (
                    voucher_blank & (debit_has_value | credit_has_value)
                )

        filled_dates = data[date_col].where(data[date_col].apply(self._is_non_empty_cell), other=pd.NA).ffill()
        has_previous_date = filled_dates.notna()
        return blank_date_mask & has_transaction_content & has_previous_date
    
    # 汇总行关键词（精确匹配，这些行不参与计算）
    # 涵盖：各种账簿软件导出时自动生成的汇总行
    # 通过词根排列组合生成：时间词 + 动作词
    SUMMARY_ROW_KEYWORDS = [
        # === 本日 ===
        '本日合计', '本日累计', '本日发生额', '本日余额', '本日结存', '本日小计',
        # === 本旬 ===
        '本旬合计', '本旬累计', '本旬发生额', '本旬余额', '本旬结存', '本旬小计',
        # === 本月 ===
        '本月合计', '本月累计', '本月发生额', '本月余额', '本月结存', '本月小计',
        # === 本季 ===
        '本季合计', '本季累计', '本季发生额', '本季余额', '本季结存', '本季小计',
        # === 本年 ===
        '本年合计', '本年累计', '本年发生额', '本年余额', '本年结存', '本年小计',
        # === 本期 ===
        '本期合计', '本期累计', '本期发生额', '本期余额', '本期结存', '本期小计',
        # === 日/月/季/年/期 + 计 ===
        '日计', '月计', '季计', '年计', '期计',
        # === 日/月/季/年 + 结 ===
        '日结', '月结', '季结', '年结',
        # === 通用汇总 ===
        '合计', '累计', '总计', '小计', '大计', '发生额', '余额', '结存',
        # === 页面相关 ===
        '本页合计', '本页累计', '本页小计', '过次页', '承前页',
        # === 期初/期末 ===
        '期初余额', '期末余额', '期初结存', '期末结存',
        '年初余额', '年末余额', '年初结存', '年末结存',
        '月初余额', '月末余额', '月初结存', '月末结存',
        # === 结转 ===
        '结转下年', '结转下期', '结转下月', '上年结转', '上期结转', '上月结转',
        '上年结余', '上期结余', '承前余额', '结转余额',
        # === 软件自动生成 ===
        '当前合计', '当前累计', '当前余额'
    ]
    
    def standardize_data(
        self,
        df: pd.DataFrame,
        mapping: Dict[str, Any],
        source_type: str,
        date_format: str = "auto",
        skiprows_offset: int = 0
    ) -> pd.DataFrame:
        """
        将 DataFrame 标准化为统一格式。
        """
        # 保留原始行号，供报表回填凭证号和未匹配明细时使用。
        working_df = df.copy()
        if '__source_row__' not in working_df.columns:
            working_df['__source_row__'] = working_df.index + 1
        if '__file_row__' not in working_df.columns:
            # pd.read_excel(skiprows=N) 跳过 N 行，第 N+1 行成为表头，
            # 第 N+2 行是第一条数据（DataFrame index 0）。
            # 因此文件行号 = index + skiprows + 2（+1 转为1起始，+1 补偿表头行）
            working_df['__file_row__'] = working_df.index + int(skiprows_offset) + 2

        # 0. 针对日记账自动执行数据清洗（跳过标题行、过滤汇总行）
        if source_type == 'journal':
            cleaner = DataCleaner(logger=self.logger)
            date_col_for_clean = mapping.get('date')
            voucher_col_for_clean = mapping.get('voucher')
            working_df = cleaner.clean_data(working_df, date_col=date_col_for_clean, voucher_col=voucher_col_for_clean)
        
        # 1. 复制副本
        data = working_df.copy()
        
        # 记录原始行号
        data['original_idx'] = data['__source_row__']
        data['original_file_row'] = data['__file_row__']
        
        # 1.5 在ffill之前，先过滤掉汇总行（本期合计、本年累计等）
        # 识别条件：日期为空 且 凭证号为空 且 摘要精确匹配关键词
        date_col = mapping.get('date')
        voucher_col = mapping.get('voucher')
        summary_col = mapping.get('summary')
        
        if summary_col and summary_col in data.columns:
            # 摘要精确匹配汇总关键词
            summary_text = data[summary_col].astype(str).str.strip()
            is_summary_keyword = summary_text.isin(self.SUMMARY_ROW_KEYWORDS)
            
            # 日期为空（包括空字符串、nan、None）
            date_is_empty = data[date_col].isna() if (date_col and date_col in data.columns) else pd.Series([True] * len(data))
            if date_col and date_col in data.columns:
                date_str = data[date_col].astype(str).str.strip()
                date_is_empty = date_is_empty | (date_str == '') | (date_str.str.lower() == 'nan')
            
            # 凭证号为空（如果有的话）
            voucher_is_empty = pd.Series([True] * len(data))
            if voucher_col and voucher_col in data.columns:
                voucher_is_empty = data[voucher_col].isna()
                voucher_str = data[voucher_col].astype(str).str.strip()
                voucher_is_empty = voucher_is_empty | (voucher_str == '') | (voucher_str.str.lower() == 'nan')
            
            # 三个条件同时满足才是汇总行
            is_summary_row = is_summary_keyword & date_is_empty & voucher_is_empty
            
            if is_summary_row.any():
                filtered_count = is_summary_row.sum()
                if self.logger:
                    self.logger(f"过滤掉 {filtered_count} 行汇总记录（本期合计/本年累计等）")
                if self.error_collector:
                    for idx in data[is_summary_row].index:
                        row_data = data.loc[idx].to_dict()
                        summary_text_val = str(row_data.get(summary_col, ''))
                        self.error_collector.record_dropped_summary_row(
                            row_idx=int(row_data.get('__file_row__', idx + 1 + int(skiprows_offset))),
                            reason=f"摘要匹配汇总关键词: {summary_text_val}",
                            original_data={k: v for k, v in row_data.items() if k not in {'__source_row__', 'original_idx'} and pd.notna(v)}
                        )
                data = data[~is_summary_row].copy()
                data['original_idx'] = data['__source_row__']
                data['original_file_row'] = data['__file_row__']
        
        all_parse_failures: List[Dict[str, Any]] = []
        
        # 2. 解析日期
        if date_col and date_col in data.columns:
            effective_date_values = data[date_col].copy()
            date_ffill_mask = self._build_date_ffill_mask(data, mapping)
            if date_ffill_mask.any():
                fill_source = effective_date_values.where(
                    effective_date_values.apply(self._is_non_empty_cell),
                    other=pd.NA
                )
                effective_date_values.loc[date_ffill_mask] = fill_source.ffill().loc[date_ffill_mask]

            data['std_date'] = effective_date_values.apply(lambda x: parse_date(x, date_format))

            invalid_date_mask = data['std_date'].isna()
            if invalid_date_mask.any() and self.error_collector:
                for idx, row in data[invalid_date_mask].iterrows():
                    original_value = row[date_col]
                    row_idx = int(row.get('original_file_row', idx + 1 + int(skiprows_offset)))
                    if pd.isna(original_value) or str(original_value).strip() == '':
                        self.error_collector.record_empty_date_row(row_idx=row_idx, source_type=source_type)
                    else:
                        self.error_collector.record_date_error(
                            row_idx=row_idx,
                            original_value=original_value,
                            source_type=source_type,
                            column=date_col
                        )
        else:
            raise ValueError(f"未找到日期列: {date_col}")

        # 3. 解析金额
        mode = mapping.get('mode', 'debit_credit')
        data['std_amount'] = Decimal('0.00')
        
        if mode == 'debit_credit':
            debit_col = mapping.get('debit')
            credit_col = mapping.get('credit')
            
            temp_debit, debit_failures = self._process_amount_column(data, debit_col, source_type)
            temp_credit, credit_failures = self._process_amount_column(data, credit_col, source_type)
            all_parse_failures.extend(debit_failures)
            all_parse_failures.extend(credit_failures)
            
            temp_debit_filled = temp_debit.fillna(Decimal('0.00'))
            temp_credit_filled = temp_credit.fillna(Decimal('0.00'))
            
            if source_type == 'bank':
                data['std_amount'] = temp_credit_filled - temp_debit_filled
            else:
                data['std_amount'] = temp_debit_filled - temp_credit_filled
            
            data['std_amount'] = data['std_amount'].apply(
                lambda x: x.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) if x is not None else None
            )
            
            debit_is_none = temp_debit.isna()
            credit_is_none = temp_credit.isna()
            # 任一侧解析失败，std_amount 设为 None
            any_parse_fail = debit_is_none | credit_is_none
            data.loc[any_parse_fail, 'std_amount'] = None
            
        elif mode in ('signed_amount', 'single_amount_with_direction'):
            amount_col = mapping.get('amount')
            direction_col = mapping.get('direction')
            
            if amount_col and amount_col in data.columns:
                if mode == 'single_amount_with_direction':
                    data['temp_amt'] = data[amount_col].apply(lambda x: clean_amount(x, allow_suffix_sign=False))
                    
                    # 区分空值和真正的0：原始值为空时，将结果设为 None 而非 Decimal('0.00')
                    amount_is_empty = data[amount_col].isna() | \
                                     (data[amount_col].astype(str).str.strip() == '') | \
                                     (data[amount_col].astype(str).str.strip().str.lower() == 'nan')
                    data.loc[amount_is_empty & (data['temp_amt'] == Decimal('0.00')), 'temp_amt'] = None
                    
                    for idx, val in data['temp_amt'].items():
                        if val is None:
                            all_parse_failures.append({
                                'row': data.loc[idx, 'original_file_row'],
                                'original_value': data.loc[idx, amount_col],
                                'source_type': source_type,
                                'column': amount_col
                            })
                            if self.error_collector:
                                self.error_collector.record_amount_error(
                                    int(data.loc[idx, 'original_file_row']),
                                    data.loc[idx, amount_col],
                                    source_type,
                                    amount_col,
                                )
                     
                    if direction_col and direction_col in data.columns:
                        direction_signs = data[direction_col].apply(
                            lambda value: direction_sign(value, source_type)
                        )
                        invalid_direction_mask = data['temp_amt'].notna() & direction_signs.isna()

                        if invalid_direction_mask.any():
                            for idx, row in data[invalid_direction_mask].iterrows():
                                original_row = int(row.get('original_file_row', idx + 1 + int(skiprows_offset)))
                                original_direction = row.get(direction_col)
                                all_parse_failures.append({
                                    'row': original_row,
                                    'original_value': original_direction,
                                    'source_type': source_type,
                                    'column': direction_col,
                                    'error_type': 'direction'
                                })
                                if self.error_collector:
                                    self.error_collector.record_direction_error(
                                        original_row,
                                        original_direction,
                                        source_type,
                                        direction_col
                                    )

                        data['std_amount'] = [
                            None if amt is None or pd.isna(sign)
                            else abs(amt) * Decimal(int(sign))
                            for amt, sign in zip(data['temp_amt'], direction_signs)
                        ]
                    else:
                        # 方向列缺失：不能静默按正数入账，必须显式报错
                        raise ValueError(
                            f"single_amount_with_direction 模式需要有效的方向列，当前映射: {direction_col}"
                        )
                else:
                    data['std_amount'] = data[amount_col].apply(
                        lambda value: parse_source_amount(value, source_type, allow_suffix_sign=True)
                    )
                    
                    # 区分空值和真正的0：原始值为空时，将结果设为 None 而非 Decimal('0.00')
                    amount_is_empty = data[amount_col].isna() | \
                                     (data[amount_col].astype(str).str.strip() == '') | \
                                     (data[amount_col].astype(str).str.strip().str.lower() == 'nan')
                    data.loc[amount_is_empty & (data['std_amount'] == Decimal('0.00')), 'std_amount'] = None
                    
                    for idx, val in data['std_amount'].items():
                        if val is None:
                            all_parse_failures.append({
                                'row': data.loc[idx, 'original_file_row'],
                                'original_value': data.loc[idx, amount_col],
                                'source_type': source_type,
                                'column': amount_col
                            })
                            if self.error_collector:
                                self.error_collector.record_amount_error(
                                    int(data.loc[idx, 'original_file_row']),
                                    data.loc[idx, amount_col],
                                    source_type,
                                    amount_col,
                                )
            else:
                raise ValueError(f"未找到金额列: {amount_col}")

        # 4. 解析摘要
        summary_col = mapping.get('summary')
        if summary_col and summary_col in data.columns:
            data['std_summary'] = data[summary_col].apply(clean_excel_string)
        else:
            data['std_summary'] = ''

        # 5. 解析余额
        balance_col = mapping.get('balance')
        if balance_col and balance_col in data.columns:
            data['std_balance'] = data[balance_col].apply(lambda x: clean_amount(x, allow_suffix_sign=False))

            # 区分空值和真正的0：原始值为空时，结果为 None 而非 Decimal('0.00')
            balance_is_empty = data[balance_col].isna() | \
                               (data[balance_col].astype(str).str.strip() == '') | \
                               (data[balance_col].astype(str).str.strip().str.lower() == 'nan')
            data.loc[balance_is_empty & (data['std_balance'] == Decimal('0.00')), 'std_balance'] = None
        else:
            data['std_balance'] = None

        # 6. 解析凭证号
        voucher_col = mapping.get('voucher')
        if voucher_col and voucher_col in data.columns:
            data['std_voucher'] = data[voucher_col].apply(clean_excel_string)
        else:
            data['std_voucher'] = ''

        # 6.5 保存多个带原列名的辅助文字字段，供本地和可选语义比较使用。
        auxiliary_columns = mapping.get('auxiliary_text_columns') or (
            [summary_col] if summary_col and summary_col in data.columns else []
        )
        auxiliary_columns = [
            column for column in auxiliary_columns
            if column in data.columns
        ]
        data['std_aux_text_fields'] = data.apply(
            lambda row: {
                str(column): clean_excel_string(row[column])
                for column in auxiliary_columns
                if self._is_non_empty_cell(row[column])
            },
            axis=1,
        )

        # 7. 构建最终 DataFrame
        result = pd.DataFrame({
            'date': data['std_date'],
            'amount': data['std_amount'],
            'summary': data['std_summary'],
            'balance': data['std_balance'],
            'voucher_no': data['std_voucher'],
            'aux_text_fields': data['std_aux_text_fields'],
            'source': source_type,
            'original_idx': data['original_idx'],
            'original_file_row': data['original_file_row']
        })
        
        # 移除无效日期行
        result = result.dropna(subset=['date'])
        
        # 记录并移除金额解析失败的行
        if result['amount'].isna().any():
            failed_count = result['amount'].isna().sum()
            result_len = len(result)
            failed_ratio = failed_count / result_len if result_len > 0 else 1.0
            
            failure_samples = []
            for failure in all_parse_failures[:10]:
                failure_samples.append(f"行{failure['row']}:'{failure['original_value']}'")
            samples_str = ", ".join(failure_samples) if failure_samples else ""
            
            if failed_ratio > 0.5:
                raise ValueError(f"{source_type} 中有 {failed_count} 行 ({failed_ratio:.1%}) 金额解析失败，请检查是否选对了金额列。失败样例: {samples_str}")
            
            if self.logger:
                self.logger(f"⚠️ 警告: {source_type} 中有 {failed_count} 行金额解析失败，已丢弃。失败样例: {samples_str}{'...' if len(all_parse_failures) > 10 else ''}")
            result = result.dropna(subset=['amount'])
        
        # 稳定排序（使用厘精度）
        result['amount_decimal'] = result['amount'].apply(PrecisionEngine.to_integer_li)
        result = result.sort_values(['date', 'amount_decimal', 'original_idx']).reset_index(drop=True)
        
        # 类型降级以节省内存
        result = _downcast_dtypes(result)
        
        return result


# ==========================================
# 4. 辅助函数
# ==========================================

def _downcast_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """降级DataFrame的数据类型以节省内存
    
    优化策略：
    1. 字符串列：如果唯一值比例<50%，转为category
    2. 整数列：降级为最小可用类型 (int8/int16/int32/int64)
    3. 浮点列：降级为 float32（注意：金额列不降级，保持Decimal）
    
    :param df: 输入DataFrame
    :return: 优化后的DataFrame
    """
    for col in df.columns:
        col_type = df[col].dtype

        if col == 'aux_text_fields':
            continue
        
        # 字符串/object列
        if col_type == 'object':
            unique_ratio = df[col].nunique() / len(df) if len(df) > 0 else 1
            AMOUNT_COLUMNS = {'date', 'amount', 'amount_decimal', 'std_amount', 
                              '借方', '贷方', '收入', '支出', '余额', 'balance'}
            if unique_ratio < 0.5 and col not in AMOUNT_COLUMNS:
                try:
                    df[col] = df[col].astype('category')
                except (TypeError, ValueError):
                    pass
        
        # 整数列降级
        elif col_type in ['int64', 'int32']:
            if df[col].min() >= 0:
                if df[col].max() < 255:
                    df[col] = df[col].astype('uint8')
                elif df[col].max() < 65535:
                    df[col] = df[col].astype('uint16')
                elif df[col].max() < 4294967295:
                    df[col] = df[col].astype('uint32')
            else:
                if df[col].min() >= -128 and df[col].max() <= 127:
                    df[col] = df[col].astype('int8')
                elif df[col].min() >= -32768 and df[col].max() <= 32767:
                    df[col] = df[col].astype('int16')
                elif df[col].min() >= -2147483648 and df[col].max() <= 2147483647:
                    df[col] = df[col].astype('int32')
    
    return df


def _gc_cleanup(stage_name: str = "", logger: Optional[Callable[[str], None]] = None) -> None:
    """执行垃圾回收并记录内存使用
    
    :param stage_name: 当前阶段名称
    :param logger: 日志函数
    """
    collected = gc.collect()
    if logger and PSUTIL_AVAILABLE:
        try:
            mem_gb = psutil.Process().memory_info().rss / (1024 ** 3)
            logger(f"🧹 [{stage_name}] GC回收 {collected} 个对象，当前内存: {mem_gb:.2f}GB")
        except (OSError, RuntimeError):
            pass
