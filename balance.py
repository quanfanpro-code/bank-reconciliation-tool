from decimal import Decimal
from typing import Any, Dict, List, Optional
import pandas as pd
from data_structures import DailyBalance, BalanceDiff, OverallControlResult
from data_loader import direction_sign, parse_source_amount
from precision_engine import PrecisionEngine
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
                    return amount * Decimal(sign)
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
                debit = debit or Decimal('0')
                credit = credit or Decimal('0')
                if source_type == 'bank':
                    return credit - debit
                if source_type == 'journal':
                    return debit - credit
                raise ValueError(f"未知数据来源: {source_type}")
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

        # 使用局部变量，避免修改实例状态，保证方法可重入
        calculated_initial = self.initial_balance
        if calculated_initial is None:
            if 'balance' not in df_sorted.columns:
                return []
            usable_balances = df_sorted['balance'].dropna()
            if not any(
                str(value).strip()
                and str(value).strip().lower() != 'nan'
                and clean_amount(value, allow_suffix_sign=False) is not None
                for value in usable_balances
            ):
                return []
            # 与 extract_initial_balance 同一套推断逻辑（容错解析并跳过空余额单元格）
            calculated_initial = self.extract_initial_balance(df_sorted)

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

            if (
                bank_balance is not None
                and journal_balance is not None
                and bank_balance == journal_balance
            ):
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
            str: 中性的业务差异类型
        """
        if bank_balance is None or journal_balance is None:
            return "数据不足"

        if bank_net != journal_net:
            return "发生额不一致"

        return "余额不一致"


def _decimal_value(value: Any) -> Optional[Decimal]:
    """将标准化表中的金额或余额安全转为 Decimal。"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, Decimal):
        return value
    return clean_amount(value, allow_suffix_sign=False)


def _date_bounds(frame: pd.DataFrame) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    if frame.empty or 'date' not in frame.columns:
        return None, None
    dates = pd.to_datetime(frame['date'], errors='coerce').dropna()
    if dates.empty:
        return None, None
    return dates.min().normalize(), dates.max().normalize()


def _amount_totals(frame: pd.DataFrame) -> tuple[Decimal, Decimal, Decimal]:
    if frame.empty or 'amount' not in frame.columns:
        zero = Decimal('0')
        return zero, zero, zero
    amounts = [
        amount
        for amount in (_decimal_value(value) for value in frame['amount'])
        if amount is not None
    ]
    income = sum((amount for amount in amounts if amount > 0), Decimal('0'))
    expense = sum((-amount for amount in amounts if amount < 0), Decimal('0'))
    return income, expense, income - expense


def check_balance_continuity(
    frame: pd.DataFrame,
    tolerance_li: int = 10,
    source: str = "",
) -> List[Dict[str, Any]]:
    """按最近有效余额和区间净额检查余额连续性，不修改输入表。"""
    if frame.empty or not {'date', 'amount', 'balance'} <= set(frame.columns):
        return []

    row_column = 'original_file_row' if 'original_file_row' in frame.columns else 'original_idx'
    sort_columns = ['date'] + ([row_column] if row_column in frame.columns else [])
    work = frame.sort_values(sort_columns, kind='stable').reset_index(drop=True).copy()
    work['date'] = pd.to_datetime(work['date'], errors='coerce').dt.normalize()
    work = work.dropna(subset=['date'])
    if work.empty:
        return []

    tolerance = PrecisionEngine.from_integer_li(tolerance_li)
    anomalies: List[Dict[str, Any]] = []
    previous_balance: Optional[Decimal] = None
    accumulated_net = Decimal('0')
    for _, row in work.iterrows():
        day = row['date']
        net = _decimal_value(row['amount'])
        balance = _decimal_value(row['balance'])
        if net is None:
            # 缺金额的区间无法可靠勾稽；下个余额只作为新基准。
            previous_balance = balance
            accumulated_net = Decimal('0')
            continue
        if balance is None:
            accumulated_net += net
            continue
        if previous_balance is not None:
            period_net = accumulated_net + net
            expected = previous_balance + period_net
            difference = abs(balance - expected)
            if difference > tolerance:
                anomalies.append({
                    '来源': source,
                    '日期': day,
                    '原文件行号': row.get('original_file_row', row.get('original_idx')),
                    '基准余额': previous_balance,
                    '区间净额': period_net,
                    '预期余额': expected,
                    '实际余额': balance,
                    '差额': difference,
                })
        previous_balance = balance
        accumulated_net = Decimal('0')
    return anomalies


def _balance_control(
    frame: pd.DataFrame,
    source: str,
    tolerance_li: int,
) -> tuple[
    Optional[Decimal], Optional[Decimal], Optional[Decimal], Optional[Decimal],
    str, tuple[Dict[str, Any], ...], tuple[str, ...]
]:
    if frame.empty or not {'date', 'amount', 'balance'} <= set(frame.columns):
        return None, None, None, None, "未实施", (), ()

    valid_balances = frame['balance'].map(_decimal_value)
    if valid_balances.dropna().empty:
        return None, None, None, None, "未实施", (), ()

    row_column = 'original_file_row' if 'original_file_row' in frame.columns else 'original_idx'
    sort_columns = ['date'] + ([row_column] if row_column in frame.columns else [])
    work = frame.sort_values(sort_columns, kind='stable').reset_index(drop=True).copy()
    parsed_balances = work['balance'].map(_decimal_value)
    last_balance_position = parsed_balances.last_valid_index()
    ending_balance = parsed_balances.loc[last_balance_position]
    initial_balance = BalanceRecalculator.extract_initial_balance(work)
    _, _, net = _amount_totals(work)
    expected_ending = initial_balance + net
    difference = abs(ending_balance - expected_ending)
    anomalies = tuple(check_balance_continuity(work, tolerance_li, source))
    tolerance = PrecisionEngine.from_integer_li(tolerance_li)
    reasons = []
    if parsed_balances.notna().sum() < 2:
        reasons.append(f"{source}仅有1个有效余额点，连续性核对未实施；倒算期初不构成独立验证")
    if difference > tolerance:
        reasons.append(f"{source}期初加净发生额与期末余额相差{difference}")
    if anomalies:
        reasons.append(f"{source}存在{len(anomalies)}处余额连续性异常")
    if last_balance_position != work.index[-1]:
        reasons.append(f"{source}最后交易行未提供余额")
    status = "疑点" if reasons else "通过"
    return (
        initial_balance,
        ending_balance,
        expected_ending,
        difference,
        status,
        anomalies,
        tuple(reasons),
    )


def build_overall_controls(
    bank: pd.DataFrame,
    journal: pd.DataFrame,
    tolerance_li: int = 10,
) -> OverallControlResult:
    """在匹配前计算双方期间、收支和余额总体控制，不修改输入表。"""
    bank_start, bank_end = _date_bounds(bank)
    journal_start, journal_end = _date_bounds(journal)
    if None in (bank_start, bank_end, journal_start, journal_end):
        period_status = "无法计算"
    elif (bank_start, bank_end) == (journal_start, journal_end):
        period_status = "通过"
    else:
        period_status = "疑点"

    bank_income, bank_expense, bank_net = _amount_totals(bank)
    journal_income, journal_expense, journal_net = _amount_totals(journal)
    if bank.empty or journal.empty:
        amount_status = "无法计算"
    elif (bank_income, bank_expense) == (journal_income, journal_expense):
        amount_status = "通过"
    else:
        amount_status = "疑点"

    bank_balance = _balance_control(bank, "银行流水", tolerance_li)
    journal_balance = _balance_control(journal, "银行日记账", tolerance_li)
    tolerance = PrecisionEngine.from_integer_li(tolerance_li)
    initial_balance_diff = (
        abs(bank_balance[0] - journal_balance[0])
        if bank_balance[0] is not None and journal_balance[0] is not None
        else None
    )
    ending_balance_diff = (
        abs(bank_balance[1] - journal_balance[1])
        if bank_balance[1] is not None and journal_balance[1] is not None
        else None
    )
    reasons = []
    if period_status != "通过":
        reasons.append(
            "双方起止日期无法确认"
            if period_status == "无法计算"
            else "双方起止日期不一致"
        )
    if amount_status != "通过":
        reasons.append(
            "双方收支金额无法确认"
            if amount_status == "无法计算"
            else "双方收入或支出合计不一致"
        )
    reasons.extend(bank_balance[6])
    reasons.extend(journal_balance[6])
    if initial_balance_diff is not None and initial_balance_diff > tolerance:
        reasons.append(f"双方期初余额相差{initial_balance_diff}")
    if ending_balance_diff is not None and ending_balance_diff > tolerance:
        reasons.append(f"双方期末余额相差{ending_balance_diff}")

    return OverallControlResult(
        bank_start_date=bank_start,
        bank_end_date=bank_end,
        journal_start_date=journal_start,
        journal_end_date=journal_end,
        period_status=period_status,
        bank_income=bank_income,
        bank_expense=bank_expense,
        bank_net=bank_net,
        journal_income=journal_income,
        journal_expense=journal_expense,
        journal_net=journal_net,
        amount_status=amount_status,
        bank_initial_balance=bank_balance[0],
        bank_ending_balance=bank_balance[1],
        bank_expected_ending_balance=bank_balance[2],
        bank_balance_diff=bank_balance[3],
        bank_balance_status=bank_balance[4],
        journal_initial_balance=journal_balance[0],
        journal_ending_balance=journal_balance[1],
        journal_expected_ending_balance=journal_balance[2],
        journal_balance_diff=journal_balance[3],
        journal_balance_status=journal_balance[4],
        initial_balance_diff=initial_balance_diff,
        ending_balance_diff=ending_balance_diff,
        continuity_anomalies=bank_balance[5] + journal_balance[5],
        scope_limited=bool(reasons),
        reasons=tuple(reasons),
    )
