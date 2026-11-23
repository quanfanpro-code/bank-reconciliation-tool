"""将最终核对结果写成五张可直接使用的工作表。"""

from collections import defaultdict
from pathlib import Path

import pandas as pd

from 排版共享 import 字母
from 排版计划 import 格样式, 应用底层主题, 新建簿计划
from 报告写入 import 渲染簿计划
from 复核事项 import build_review_items


def _text(value):
    return str(value).strip() if value is not None and pd.notna(value) else ""


def _records(frame, indices):
    for index in indices:
        row = frame.loc[index]
        fields = row.get("aux_text_fields", {})
        fields = fields if isinstance(fields, dict) else {}
        party = "、".join(dict.fromkeys(
            _text(v) for k, v in fields.items() if any(s in k for s in ("对方", "收款人", "付款人")) and _text(v)
        ))
        note = "；".join(dict.fromkeys(
            _text(v) for k, v in fields.items() if any(s in k for s in ("备注", "用途", "附言")) and _text(v)
        ))
        yield row, party, note


def _brief(rows):
    values = []
    for row, party, note in rows[:3]:
        voucher = _text(row.get("voucher_no"))
        context = "，".join(filter(None, (party, _text(row.get("summary")), note, f"凭证{voucher}" if voucher else "")))
        action = "收款" if row["amount"] >= 0 else "付款"
        values.append(f"{row['date']:%m-%d} {action}{abs(row['amount']):,.2f}元 {context}".strip())
    if len(rows) > 3:
        values.append(f"其余{len(rows) - 3}笔见交易明细")
    return "\n".join(values)


def _periods(frame, key, ending, notes):
    original = key
    frame = frame.rename(columns=lambda col: col.replace("日记账-", "账上").replace("银行-", "银行").replace("期末余额", ending))
    balance_labels = [f"银行{ending}", f"账上{ending}", "余额差额"]
    columns = [key, "银行收入金额", "账上收入金额", "收入金额差额", "银行支出金额", "账上支出金额", "支出金额差额",
               *balance_labels, "收入笔数差额", "支出笔数差额"]
    result = frame.reindex(columns=columns).copy()
    if key == "月份":
        result[key] = result[key].astype(str)
    result["说明"] = result[original].map(lambda value: "；".join(notes.get(value, ())))
    return result


def write_reconciliation_report(reporter, output_path, config, bank_path, journal_path):
    reporter._prepare_initial_balance()
    daily, monthly = reporter._build_daily_and_monthly_tables()
    items = sorted(build_review_items(reporter.matcher), key=lambda item: (item["首次日期"], item["事项编号"]))
    paths = {"银行": bank_path, "账上": journal_path}
    origins = {}
    for source in getattr(reporter.precheck_report, "source_info", ()):
        side = "银行" if source.get("来源") == "银行流水" else "账上"
        origins[side] = " · ".join(filter(None, (Path(source["文件路径"]).name, source.get("工作表"))))
    daily_notes, monthly_notes = defaultdict(list), defaultdict(list)
    pending, details = [], []
    for number, item in enumerate(items, 1):
        label = f"{number:03d}"
        bank = list(_records(reporter.matcher.bank, item["银行索引"]))
        journal = list(_records(reporter.matcher.journal, item["日记账索引"]))
        mode = item["核对方式"]
        confirmed = mode == "自动确认"
        action = {"人工全查": "逐项核对", "人工抽样": "抽查"}.get(mode)
        income = item["银行收入"] - item["日记账收入"]
        expense = item["银行支出"] - item["日记账支出"]
        dates = sorted({row["date"].normalize() for row, _, _ in bank + journal})
        months = sorted({day.strftime("%Y-%m") for day in dates})
        if confirmed:
            explanation = "银行与账上金额一致"
            if len(bank) != len(journal):
                explanation = f"银行{len(bank)}笔对应账上{len(journal)}笔，合计金额一致"
            if len(dates) > 1:
                explanation += "，记账日期不同"
                for day in dates:
                    daily_notes[day].append(f"{label} 已对上，日期不同")
                if len(months) > 1:
                    for month in months:
                        monthly_notes[month].append(f"{label} 已对上，入账月份不同")
        elif item["程序结论"] in {"账内更正已闭合", "银行冲正已闭合"}:
            explanation = "原交易已冲销，净额为零"
        elif not bank or not journal:
            explanation = "银行有记录，账上未找到对应" if bank else "账上有记录，银行未找到对应"
        elif income or expense:
            explanation = "双方记录之间仍有收付金额差额"
        else:
            explanation = "双方金额相同，交易对应尚未确定"
        if action:
            pending.append((label, action, explanation, _brief(bank), _brief(journal), income, expense, "", "", "查看明细"))
        for side, entries in (("银行", bank), ("账上", journal)):
            origin = origins.get(side) or (Path(paths[side]).name if paths[side] else side)
            for row, party, note in entries:
                details.append((label, side, row["date"], max(row["amount"], 0), max(-row["amount"], 0),
                                row.get("balance"), party, row.get("summary", ""), note, _text(row.get("voucher_no")),
                                explanation, f"{origin} · 第{row.get('original_file_row', row.get('original_idx', ''))}行"))

    months_with_difference = sum(
        any(pd.notna(value) and value != 0 for value in values)
        for values in monthly.reindex(columns=["收入金额差额", "支出金额差额", "余额差额"]).itertuples(index=False, name=None)
    )
    dates = sorted(set(reporter.matcher.bank["date"]).union(reporter.matcher.journal["date"]))
    message = f"{months_with_difference} 个月有收支或余额差额，见月度核对" if months_with_difference else "各月收入及支出金额一致"
    if not months_with_difference and "余额差额" in monthly and not monthly.empty and monthly["余额差额"].notna().all():
        message = "各月收入、支出及月末余额一致"
    filename = Path(bank_path).stem if bank_path else ""
    identity = filename.rsplit("-", 2)
    overview = [
        ("公司", identity[0] if len(identity) == 3 else ""), ("银行及账号", " · ".join(identity[1:]) if len(identity) == 3 else ""),
        ("银行流水", Path(bank_path).name if bank_path else ""), ("银行日记账", Path(journal_path).name if journal_path else ""),
        ("交易日期", f"{dates[0]:%Y-%m-%d} 至 {dates[-1]:%Y-%m-%d}" if dates else "无交易"),
        ("读入笔数", f"银行 {len(reporter.matcher.bank)} 笔，账上 {len(reporter.matcher.journal)} 笔"),
        ("月度核对结果", message), ("需要人工核对", len(pending)),
        ("明显微小错报临界值", config.clearly_trivial_threshold), ("实际执行重要性水平", config.performance_materiality),
    ]
    if reporter.error_collector and reporter.error_collector.has_errors():
        for error in reporter.error_collector.get_all_errors():
            overview.append(("未读入的记录", f"{error.get('source_type', '')} 第{error.get('row', '')}行：{error.get('reason', '')}"))
    pending_frame = pd.DataFrame(pending or [("", "", "本次无需人工核对", "", "", None, None, "", "", "")], columns=[
        "编号", "处理方式", "要查什么", "银行记录", "账上记录", "收入差额", "支出差额", "核查结论", "备注", "交易明细"])
    tables = {
        "核对总览": pd.DataFrame(overview, columns=["项目", "结果"]),
        "月度核对": _periods(monthly, "月份", "月末余额", monthly_notes),
        "日度核对": _periods(daily, "日期", "日末余额", daily_notes),
        "待核对": pending_frame,
        "交易明细": pd.DataFrame(details, columns=["编号", "来源", "日期", "收入金额", "支出金额", "余额", "对方", "摘要", "银行备注", "凭证号", "核对结果", "原表位置"]),
    }
    plan = 新建簿计划({name: reporter._safe_table(table) for name, table in tables.items()})
    for sheet in plan.表:
        应用底层主题(sheet, 首列留白=False)
        sheet.冻结 = "A2"
        if sheet.名称 != "核对总览" and len(sheet.表):
            sheet.筛选 = f"A1:{字母(len(sheet.表.columns))}{len(sheet.表) + 1}"
    todo = plan.取表("待核对")
    if pending:
        todo.表["交易明细"] = [
            f'=HYPERLINK("#\'交易明细\'!A"&(MATCH(A{row},\'交易明细\'!$A$2:$A${len(details) + 1},0)+1),"查看明细")'
            for row in range(2, len(pending) + 2)
        ]
    for row in range(2, len(pending) + 2):
        for column in (8, 9):
            todo.盖格(row, column, 格样式(填充色="FFF2CC", 填充类型="solid"))
        todo.行计划取(row).高 = 68
    渲染簿计划(plan, output_path)
