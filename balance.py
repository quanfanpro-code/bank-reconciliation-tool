from decimal import Decimal
from typing import Any, Dict, List, Optional
import pandas as pd
from data_structures import DailyBalance, BalanceDiff
from data_loader import direction_sign, parse_source_amount
from utils import clean_amount


class BalanceRecalculator:
    """余额重算器 - 按日期重新计算每日余额"""

    def __init__(
        self,
        initial_balance: Optional[Decimal] = None,
    ):
        self.initial_balance = initial_balance

    @staticmethod
    def extract_initial_balance(
        df: pd.DataFrame,
        mapping: Optional[Dict[str, Any]] = None,
        source_type: str = "journal",
    ) -> Decimal:
        """提取期初余额。

        列名解析：优先使用 mapping（用户校验过的列映射），
        mapping 未提供或列不存在时回退到内置列名猜测（精确匹配）。
        金额取值一律走 clean_amount 容错解析；解析失败的行跳过，不抛异常。
        确实无法确定期初余额时返回 Decimal('0')。
        """
        if df.empty:
            return Decimal('0')

        mapping = mapping or {}

        # 内置列名猜测（精确匹配，避免误匹配）
        date_col = summary_col = balance_col = amount_col = None
        debit_col = credit_col = None
        for col in df.columns:
            col_lower = str(col).lower().strip()
            col_stripped = str(col).strip()
            if col_stripped in ('date', '日期', 'std_date') or col_lower == 'date':
                date_col = date_col or col
            elif col_stripped in ('summary', '摘要', 'std_summary') or col_lower == 'summary':
                summary_col = summary_col or col
            elif col_stripped in ('balance', '余额', 'std_balance') or col_lower == 'balance':
                balance_col = balance_col or col
            elif col == 'amount':
                amount_col = amount_col or col
            if 'debit' in col_lower or '借方' in col_stripped:
                debit_col = debit_col or col
            elif 'credit' in col_lower or '贷方' in col_stripped:
                credit_col = credit_col or col

        # mapping 指定的列优先（列必须真实存在）
        def _pick(key: str, guessed: Optional[str]) -> Optional[str]:
            col = mapping.get(key)
            return col if (col and col in df.columns) else guessed

        date_col = _pick('date', date_col)
        summary_col = _pick('summary', summary_col)
        balance_col = _pick('balance', balance_col)
        amount_col = _pick('amount', amount_col)
        debit_col = _pick('debit', debit_col)
        credit_col = _pick('credit', credit_col)
        direction_col = _pick('direction', None)
        mode = mapping.get('mode', 'debit_credit')

        if not balance_col:
            return Decimal('0')
        if not date_col:
            date_col = list(df.columns)[0]

        work = df.copy()
        work['__parsed_date__'] = pd.to_datetime(work[date_col], errors='coerce')
        sort_cols = ['__parsed_date__'] + (['original_idx'] if 'original_idx' in work.columns else [])
        work = work.sort_values(sort_cols, kind='stable', na_position='last')

        def _parse(val) -> Optional[Decimal]:
            """空单元格 -> None（区别于真 0），其余走 clean_amount 容错解析。"""
            if pd.isna(val):
                return None
            text = str(val).strip()
            if text == '' or text.lower() == 'nan':
                return None
            return clean_amount(val, allow_suffix_sign=False)

        # 1) 期初标记行（"上期结转"/"期初"等，日期通常为空）
        if summary_col:
            initial_keywords = ['上期结转', '期初余额', '年初余额', '月初余额', '期初', '结转']
            kw_pattern = '|'.join(initial_keywords)
            kw_mask = (work['__parsed_date__'].isna()
                       & work[summary_col].astype(str).str.contains(kw_pattern, na=False))
            for val in work.loc[kw_mask, balance_col]:
                parsed = _parse(val)
                if parsed is not None:
                    return parsed

        # 2) 首个有效余额行回补推算：
        #    期初 = 该行余额 - 截至该行（含）的累计净额
        #    注意：借贷分列仅用于日记账原始数据（借方-贷方为企业视角净额）
        def _row_net(row) -> Optional[Decimal]:
            if amount_col:
                if mode == 'single_amount_with_direction':
                    if not direction_col:
                        return None
                    amount = _parse(row[amount_col])
                    sign = direction_sign(row[direction_col], source_type)
                    if amount is None or sign is None:
                        return None
                    return abs(amount) * Decimal(sign)
                if mode == 'signed_amount':
                    return parse_source_amount(
                        row[amount_col],
                        source_type,
                        allow_suffix_sign=True,
                    )
                return _parse(row[amount_col])
            if debit_col and credit_col:
                debit = _parse(row[debit_col])
                credit = _parse(row[credit_col])
                if debit is None and credit is None:
                    return None
                return (debit or Decimal('0')) - (credit or Decimal('0'))
            return None

        cumulative = Decimal('0')
        for _, row in work.iterrows():
            net = _row_net(row)
            if net is not None:
                cumulative += net
            parsed_balance = _parse(row[balance_col])
            if parsed_balance is not None:
                return parsed_balance - cumulative

        return Decimal('0')

    def recalculate(self, df: pd.DataFrame) -> List[DailyBalance]:
        if df.empty or 'amount' not in df.columns:
            return []

        df_sorted = df.sort_values('date').copy()
        df_sorted['date'] = pd.to_datetime(df_sorted['date'])

        daily_net = df_sorted.groupby('date')['amount'].sum()
        daily_net_dict = {pd.Timestamp(k).date(): Decimal(str(v)) for k, v in daily_net.items()}

        all_dates = pd.date_range(
            start=df_sorted['date'].min(),
            end=df_sorted['date'].max(),
            freq='D'
        )

        first_date = all_dates[0].date()
        first_day_net = daily_net_dict.get(first_date, Decimal('0'))

        # 使用局部变量，避免修改实例状态，保证方法可重入
        calculated_initial = self.initial_balance
        if calculated_initial is None and 'balance' in df_sorted.columns:
            # 与 extract_initial_balance 同一套推断逻辑（容错解析并跳过空余额单元格）
            calculated_initial = self.extract_initial_balance(df_sorted)

        if calculated_initial is None:
            calculated_initial = Decimal('0')

        results = []
        prev_balance = calculated_initial

        for date in all_dates:
            date_key = date.date()
            net = daily_net_dict.get(date_key, Decimal('0'))

            income = net if net > 0 else Decimal('0')
            expense = abs(net) if net < 0 else Decimal('0')

            balance = prev_balance + net

            results.append(DailyBalance(
                date=pd.Timestamp(date),
                income=income,
                expense=expense,
                net=net,
                balance=balance,
                prev_balance=prev_balance
            ))

            prev_balance = balance

        return results

class BalanceReconciler:
    """余额核对器 - 比较银行与日记账的重算余额差异"""

    def __init__(
        self,
        bank_balances: List[DailyBalance],
        journal_balances: List[DailyBalance]
    ):
        """
        初始化余额核对器

        参数:
            bank_balances: 银行流水的重算每日余额列表
            journal_balances: 日记账的重算每日余额列表
        """
        self.bank_balances = bank_balances
        self.journal_balances = journal_balances

    def generate_diff_report(self) -> List[BalanceDiff]:
        """
        生成余额差异报告

        返回:
            List[BalanceDiff]: 差异报告列表，按日期排序
        """
        bank_daily = {b.date: b for b in self.bank_balances}
        journal_daily = {b.date: b for b in self.journal_balances}

        all_dates = sorted(set(bank_daily.keys()) | set(journal_daily.keys()))

        report = []
        for date in all_dates:
            bank_info = bank_daily.get(date)
            journal_info = journal_daily.get(date)

            bank_balance = bank_info.balance if bank_info else None
            journal_balance = journal_info.balance if journal_info else None
            bank_net = bank_info.net if bank_info else Decimal('0.00')
            journal_net = journal_info.net if journal_info else Decimal('0.00')

            # 双方都无数据（None）时跳过，但余额恰好为 0 不跳过
            if bank_balance is None and journal_balance is None:
                continue

            diff_type = self.classify_diff(
                bank_balance=bank_balance,
                journal_balance=journal_balance,
                bank_net=bank_net,
                journal_net=journal_net
            )

            if bank_balance is None:
                bank_balance = Decimal('0.00')
            if journal_balance is None:
                journal_balance = Decimal('0.00')

            diff = bank_balance - journal_balance

            report.append(BalanceDiff(
                date=pd.Timestamp(date),
                bank_balance=bank_balance,
                journal_balance=journal_balance,
                diff=diff,
                diff_type=diff_type
            ))

        return report

    def classify_diff(
        self,
        bank_balance: Optional[Decimal],
        journal_balance: Optional[Decimal],
        bank_net: Decimal,
        journal_net: Decimal
    ) -> str:
        """
        分类差异类型

        参数:
            bank_balance: 银行余额（可能为空）
            journal_balance: 日记账余额（可能为空）
            bank_net: 银行当日净额
            journal_net: 日记账当日净额

        返回:
            str: 差异类型 ('时间序错误'/'金额错误'/'缺失记录')
        """
        if bank_balance is None or journal_balance is None:
            return "缺失记录"

        if bank_net != journal_net:
            return "金额错误"

        return "时间序错误"
