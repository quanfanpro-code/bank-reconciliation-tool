"""把有日期的余额、逐笔覆盖和收付差额分别勾稽。"""
from datetime import date
from decimal import Decimal, InvalidOperation

import pandas as pd

from utils import round_decimal


def parse_coverage_evidence(data):
    """按银行或账上来源及完整文件名绑定凭据。"""
    if not isinstance(data, dict) or data.keys() - {"bank", "journal"}:
        raise ValueError("覆盖凭据须按 bank 和 journal 区分来源")
    return {side: _parse_files(files) for side, files in data.items()}


def _parse_files(data):
    if not isinstance(data, dict):
        raise ValueError("覆盖凭据必须按输入文件名提供")
    result = {}
    for filename, raw in data.items():
        if not filename or "/" in filename or chr(92) in filename or not isinstance(raw, dict):
            raise ValueError("覆盖凭据必须使用输入文件的完整文件名")
        balances, periods = {}, []
        for point in raw.get("balances", []):
            day = _day(point["date"])
            source = _source(point["source"])
            try:
                amount = Decimal(str(point["amount"]))
            except InvalidOperation as exc:
                raise ValueError("余额凭据的金额无效") from exc
            if not amount.is_finite():
                raise ValueError("余额凭据的金额无效")
            if day in balances and balances[day][0] != amount:
                raise ValueError(f"{filename}在{day:%Y-%m-%d}有两条不同的银行余额凭据")
            balances[day] = (amount, source)
        for period in raw.get("coverage", []):
            start, end = _day(period["start"]), _day(period["end"])
            kind, source = period["kind"], _source(period["source"])
            if start > end or kind not in {"完整", "无交易", "不完整"}:
                raise ValueError("覆盖范围或类别有误")
            periods.append((start, end, kind, source))
        result[filename] = {"balances": balances, "periods": periods}
    return result


def _day(value):
    return pd.Timestamp(date.fromisoformat(value))


def _source(value):
    if not isinstance(value, str) or not value.strip():
        raise ValueError("覆盖凭据缺少原件出处")
    return value.strip()


def coverage_dates(evidence):
    for side in evidence.values():
        yield from side.get("balances", {})
        for start, end, _, _ in side.get("periods", []):
            yield start
            yield end


def uncovered_dates(days, missing_months, evidence):
    missing = days.dt.to_period("M").isin(missing_months)
    for start, end, kind, _ in evidence.get("periods", []):
        if kind == "不完整":
            missing |= days.between(start, end)
    for start, end, kind, _ in evidence.get("periods", []):
        if kind != "不完整":
            missing &= ~days.between(start, end)
    return missing


def _subtract(left, right):
    return [round_decimal(Decimal(str(a)) - Decimal(str(b))) if pd.notna(a) and pd.notna(b) else None
            for a, b in zip(left, right)]


def apply_coverage(daily, monthly, matcher, daily_notes, monthly_notes):
    if daily.empty or not hasattr(matcher, "uncovered_months"):
        return
    evidence = getattr(matcher, "coverage_evidence", {})
    days = pd.to_datetime(daily["日期"]).dt.normalize()
    periods = days.dt.to_period("M")
    days_by_month = {month: periods.eq(month) for month in periods.unique()}
    month_keys = monthly["月份"].astype(str)
    for side, prefix, label in (("bank", "银行", "银行"), ("journal", "日记账", "账上")):
        facts = evidence.get(side, {})
        missing = uncovered_dates(days, matcher.uncovered_months.get(side, ()), facts)
        amount_columns = [f"{prefix}-{key}" for key in ("收入金额", "支出金额", "收入笔数", "支出笔数")]
        daily.loc[missing, amount_columns] = None
        for day in days[missing]:
            daily_notes[day].append(f"{label}逐笔明细未覆盖")
        for month in periods[missing].unique():
            monthly.loc[month_keys.eq(str(month)), amount_columns] = None
            monthly_notes[str(month)].append(f"{label}逐笔明细未覆盖，整月发生额无法核对")
        balance_column = f"{prefix}-期末余额"
        month_column = f"{prefix}-月末余额"
        if balance_column not in daily:
            daily[balance_column] = None
        if month_column not in monthly:
            monthly[month_column] = None
        # 明细缺口切断余额结转，从下一条有日期的原始余额重新起算。
        previous = None
        frame = matcher.bank if side == "bank" else matcher.journal
        original_days = set(frame["date"].dt.normalize())
        for index, day in days.items():
            point = facts.get("balances", {}).get(day)
            if point:
                amount, source = point
                original = daily.at[index, balance_column]
                if day in original_days and pd.notna(original) and Decimal(str(original)) != amount:
                    daily_notes[day].append(
                        f"{label}余额凭据与交易记录日末余额相差{abs(amount - Decimal(str(original))):,.2f}元"
                    )
                previous = round_decimal(amount)
                daily_notes[day].append(f"{label}余额来源 {source}")
            elif missing.at[index]:
                previous = None
            elif day in original_days:
                previous = daily.at[index, balance_column]
            daily.at[index, balance_column] = previous
        for month, mask in days_by_month.items():
            last = daily.loc[mask].iloc[-1]
            if missing[mask].any() and last["日期"].day != month.days_in_month:
                value = None
            else:
                value = last[balance_column]
            monthly.loc[month_keys.eq(str(month)), month_column] = value
        for start, end, kind, source in facts.get("periods", []):
            if kind == "无交易":
                for day in days[days.between(start, end)]:
                    daily_notes[day].append(f"{label}查询确认无交易 {source}")

    for frame, balance_key in ((daily, "期末余额"), (monthly, "月末余额")):
        for key in ("收入金额", "支出金额", "收入笔数", "支出笔数", balance_key):
            difference = "余额差额" if key == balance_key else f"{key}差额"
            frame[difference] = _subtract(frame[f"银行-{key}"], frame[f"日记账-{key}"])
        for row in frame.itertuples(index=False, name=None):
            record = dict(zip(frame.columns, row))
            period = record["日期"] if frame is daily else str(record["月份"])
            notes = daily_notes if frame is daily else monthly_notes
            balance_diff = record["余额差额"]
            amounts = (record["收入金额差额"], record["支出金额差额"])
            if pd.notna(balance_diff) and balance_diff:
                notes[period].append(f"{'日末' if frame is daily else '月末'}余额相差{abs(balance_diff):,.2f}元")
            elif pd.notna(balance_diff):
                if any(pd.isna(value) for value in amounts):
                    notes[period].append("余额相符，收付无法完整核对")
                elif any(amounts):
                    notes[period].append("余额相符，仍有收付差额")
    for month, mask in days_by_month.items():
        mismatches = [abs(value) for value in daily.loc[mask, "余额差额"] if pd.notna(value) and value]
        if mismatches:
            monthly_notes[str(month)].append(f"日末余额有{len(mismatches)}天不同，最大差额{max(mismatches):,.2f}元")
