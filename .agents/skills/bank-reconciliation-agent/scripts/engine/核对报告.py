"""将最终核对结果写成可直接使用的工作表。"""

from collections import defaultdict
from pathlib import Path

import pandas as pd

from Excel验证 import 明细跳转公式
from 排版共享 import 字母
from 排版计划 import 格样式, 应用底层主题, 新建簿计划
from 报告写入 import 渲染簿计划
from 复核事项 import build_review_items
from 核对期间 import apply_coverage
from utils import parse_date


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


def _brief(rows, first_row, label, detail_tab):
    values = []
    for row, party, note in rows[:3]:
        voucher = _text(row.get("voucher_no"))
        context = "，".join(filter(None, (party, _text(row.get("summary")), note, f"凭证{voucher}" if voucher else "")))
        action = "收款" if row["amount"] >= 0 else "付款"
        values.append(f"{row['date']:%m-%d} {action}{abs(row['amount']):,.2f}元 {context}".strip())
    if len(rows) > 3:
        values.append(f"另{len(rows) - 3}笔见{detail_tab}第{first_row + 3}–{first_row + len(rows) - 1}行（编号{label}）")
    return "\n".join(values)


def _difference_text(income, expense):
    return "，".join(
        f"{'银行' if value > 0 else '账上'}{direction}多 {abs(value):,.2f} 元"
        for direction, value in (("收入", income), ("支出", expense)) if value
    )


def _explain_fee_differences(monthly, matcher, notes):
    """手续费公式已经说清的净额，从月度未解释差额里剔除。"""
    from decimal import Decimal

    from precision_engine import PrecisionEngine

    if monthly is None or getattr(monthly, "empty", True):
        return
    explained = {}
    for candidate in getattr(matcher, "selected_candidates", []):
        if getattr(candidate, "match_type", "") != "fee_net":
            continue
        if candidate.processing_status.value not in {"自动确认", "整组勾稽一致"}:
            continue
        months = {
            pd.Timestamp(day).strftime("%Y-%m")
            for day in (*candidate.bank_dates, *candidate.journal_dates)
        }
        if len(months) != 1:
            continue
        month = months.pop()
        metrics = candidate.metrics
        income = PrecisionEngine.from_integer_li(metrics.bank_income_li - metrics.journal_income_li)
        expense = PrecisionEngine.from_integer_li(metrics.bank_expense_li - metrics.journal_expense_li)
        fee = PrecisionEngine.from_integer_li(int(candidate.evidence.get("fee_amount_li") or 0))
        bucket = explained.setdefault(month, [Decimal("0"), Decimal("0"), Decimal("0")])
        bucket[0] += income
        bucket[1] += expense
        bucket[2] += fee
    for month, (income, expense, fee) in explained.items():
        mask = monthly["月份"].astype(str) == month
        if not bool(mask.any()):
            continue
        for column, delta in (("收入金额差额", income), ("支出金额差额", expense)):
            monthly.loc[mask, column] = [
                float(Decimal(str(value)) - delta) if pd.notna(value) else value
                for value in monthly.loc[mask, column]
            ]
        notes[month].append(f"手续费净额已说明 {fee:,.2f} 元，这项差额不再作为未解释的收付差额")


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
    automatic = getattr(reporter.matcher, "skill_automatic", False)
    report_tab, detail_tab = ("差异说明", "差异明细") if automatic else ("待核对", "待核对明细")
    chosen = {candidate.final_match_id for candidate in reporter.matcher.selected_candidates
              if candidate.metrics.total_diff_li == 0 and not candidate.evidence.get("selection_ineligible")}
    duplicates = defaultdict(list)
    for clue in getattr(reporter.matcher, "business_clues", []):
        if "强重复" in clue.clue_type:
            for index in clue.row_idxs:
                duplicates[clue.source, index].append(clue.reason)
    agent_notes = {
        candidate.final_match_id: candidate.llm_decision.reason
        for candidate in reporter.matcher.selected_candidates
        if candidate.evidence.get("agent_confirmed") and candidate.llm_decision
    }
    skill_groups = {c.final_match_id for c in reporter.matcher.selected_candidates if c.match_type == "skill_closed_cover"}
    paths = {"银行": bank_path, "账上": journal_path}
    origins = {}
    for source in getattr(reporter.precheck_report, "source_info", ()):
        side = "银行" if source.get("来源") == "银行流水" else "账上"
        origins[side] = " · ".join(filter(None, (Path(source["文件路径"]).name, source.get("工作表"))))
    daily_notes, monthly_notes = defaultdict(list), defaultdict(list)
    apply_coverage(daily, monthly, reporter.matcher, daily_notes, monthly_notes)
    scope_conflicts = getattr(reporter.matcher, "scope_conflicts", ())
    scope_message = "，".join(
        f"{item.name}：{item.comparison}" + (f"（银行 {item.bank_result}，账上 {item.journal_result}）"
                                             if item.bank_result and item.journal_result else "")
        for item in scope_conflicts
    )
    if scope_conflicts:
        prefixes = {"银行-": "银行", "日记账-": "账上"}
        empty_sides = getattr(reporter.matcher, "empty_unverified_sides", ())
        for table in (daily, monthly):
            for column in table.columns:
                if "差额" in column or any(side in empty_sides and column.startswith(prefix) for prefix, side in prefixes.items()):
                    table[column] = None
    unread = [row for row in getattr(reporter.precheck_report, "population_rows", ())
              if row.get("唯一处置类别") in {"解析异常", "其他排除"}]
    for row in unread:
        source = row["来源"]
        day = parse_date(row.get("原日期值"))
        note = f"{source}{row.get('工作表', '')}第{row['原文件行号']}行未计入。{row['处置原因']}"
        if day is not None:
            daily_notes[day].append(note)
        for table, frequency, date_column, notes in ((daily, "D", "日期", daily_notes), (monthly, "M", "月份", monthly_notes)):
            mask = pd.Series(True, index=table.index)
            if day is not None:
                mask = pd.to_datetime(table[date_column]).dt.to_period(frequency).eq(day.to_period(frequency))
            for field in ("收入金额差额", "支出金额差额", "收入笔数差额", "支出笔数差额"):
                if field in table:
                    table[field] = table[field].mask(mask)
            if frequency == "M":
                for month in table.loc[mask, "月份"]:
                    notes[str(month)].append(note)
    for entry in getattr(reporter.matcher.overall_control, "continuity_anomalies", ()):
        day = pd.Timestamp(entry["日期"]).normalize()
        note = (f"{entry['来源']} 第{entry['原文件行号']}行按当日收付应余 {entry['预期余额']:,.2f} 元，"
                f"表列 {entry['实际余额']:,.2f} 元，相差 {entry['差额']:,.2f} 元")
        daily_notes[day].append(note)
        monthly_notes[day.strftime("%Y-%m")].append(f"{day:%m-%d} {note}")
    pending, details, review_details, review_groups = [], [], [], []
    for number, item in enumerate(items, 1):
        label = f"{number:03d}"
        bank = list(_records(reporter.matcher.bank, item["银行索引"]))
        journal = list(_records(reporter.matcher.journal, item["日记账索引"]))
        mode = item["核对方式"]
        confirmed = mode == "自动确认"
        action = {"人工全查": "逐项核对", "人工抽样": "抽查", "范围说明": "范围说明", "差异说明": "差异说明"}.get(mode)
        income = item["银行收入"] - item["日记账收入"]
        expense = item["银行支出"] - item["日记账支出"]
        dates = sorted({row["date"].normalize() for row, _, _ in bank + journal})
        months = sorted({day.strftime("%Y-%m") for day in dates})
        if item.get("明细覆盖不足"):
            month = dates[0]
            missing, observed, count = ("银行", "账上", len(journal)) if journal else ("账上", "银行", len(bank))
            record_name = "流水" if journal else "日记账"
            explanation = f"{month.year}年{month.month}月{missing}缺这些日期的{record_name}，{observed}有{count}笔。收付差额无法计算。"
            income = expense = None
        elif confirmed:
            explanation = "银行与账上金额一致"
            if len(bank) != len(journal):
                explanation = f"银行{len(bank)}笔对应账上{len(journal)}笔，合计金额一致"
            elif len(bank) > 1 and item["匹配ID"] in skill_groups:
                explanation = f"银行{len(bank)}笔与账上{len(journal)}笔整组核对，收入和支出分别一致"
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
            explanation = scope_message if scope_conflicts else (
                "银行有记录，账上未找到对应" if bank else "账上有记录，银行未找到对应")
            if scope_conflicts:
                income = expense = None
        elif income or expense:
            explanation = "双方记录之间仍有收付金额差额"
        elif automatic and item["匹配ID"] in chosen:
            explanation = "已选定对应，银行与账上收付分别一致"
            action = None
        else:
            explanation = "双方金额相同，交易对应尚未确定"
        if (action or automatic and not confirmed and item["匹配ID"] in chosen) and not item.get("明细覆盖不足"):
            reasons = [
                reason for reason in item["判断依据"].split("；")
                if reason and not reason.startswith("综合可信度") and reason != "本次匹配未建立对应关系"
            ]
            if reasons:
                explanation += "。具体问题：" + "；".join(reasons)
            if item["入选原因"] and not automatic:
                explanation += "。入选原因：" + item["入选原因"]
        if item["事项编号"] in agent_notes:
            explanation += "。Agent 判断依据：" + agent_notes[item["事项编号"]]
        first_detail_row = len(details) + 2
        for side, entries in (("银行", bank), ("账上", journal)):
            origin = origins.get(side) or (Path(paths[side]).name if paths[side] else side)
            for row, party, note in entries:
                source = "bank" if side == "银行" else "journal"
                clues = duplicates.get((source, row.name), ())
                result = "。".join((explanation, *clues))
                page = row.get("original_sheet")
                location = f"{Path(paths[side]).name} · {page}" if paths[side] and page else origin
                details.append((label, side, row["date"], row.get("income", max(row["amount"], 0)), row.get("expense", max(-row["amount"], 0)),
                                row.get("balance"), party, row.get("summary", ""), note, _text(row.get("voucher_no")),
                                result, f"{location} · 第{row.get('original_file_row', row.get('original_idx', ''))}行"))
        if action:
            first_review_row = len(review_details) + 2
            count = len(bank) + len(journal)
            pending.append((label, action, explanation, _brief(bank, first_review_row, label, detail_tab),
                            _brief(journal, first_review_row + len(bank), label, detail_tab), income, expense, "", "", ""))
            review_groups.append((action, first_detail_row, first_review_row, count))
            review_details.extend(details[first_detail_row - 2:])
            if automatic and (difference := _difference_text(income, expense)):
                sources = "、".join(row[-1] for row in details[first_detail_row - 2:])
                text = f"{difference}。原行 {sources}"
                if len(dates) == 1:
                    daily_notes[dates[0]].append(text)
                if len(months) == 1:
                    monthly_notes[months[0]].append(f"{dates[0]:%m月%d日} {text}")

    _explain_fee_differences(monthly, reporter.matcher, monthly_notes)
    months_with_difference = sum(
        any(pd.notna(value) and value != 0 for value in values)
        for values in monthly.reindex(columns=["收入金额差额", "支出金额差额", "余额差额"]).itertuples(index=False, name=None)
    )
    dates = sorted(set(reporter.matcher.bank["date"]).union(reporter.matcher.journal["date"]))
    message = f"{months_with_difference} 个月有收支或余额差额，见月度核对" if months_with_difference else "各月收入及支出金额一致"
    if not months_with_difference and "余额差额" in monthly and not monthly.empty and monthly["余额差额"].notna().all():
        message = "各月收入、支出及月末余额一致"
    if hasattr(reporter.matcher, "uncovered_months"):
        gaps = monthly[["收入金额差额", "支出金额差额"]].isna().any(axis=1)
        if gaps.any():
            compared = monthly.loc[~gaps]
            unequal = sum(any(pd.notna(v) and v != 0 for v in row)
                          for row in compared[["收入金额差额", "支出金额差额"]].itertuples(index=False, name=None))
            message = f"{int(gaps.sum())} 个月收付无法完整核对，已覆盖月份有 {unequal} 个月收付有差额"
        if "余额差额" in daily:
            balance_days = [abs(v) for v in daily["余额差额"] if pd.notna(v) and v]
            if balance_days:
                message += f"。日末余额有 {len(balance_days)} 天不同，最大差 {max(balance_days):,.2f} 元"
    filename = Path(bank_path or journal_path).stem if bank_path or journal_path else ""
    identity = filename.rsplit("-", 2)
    accounts = {str(value) for frame in (reporter.matcher.bank, reporter.matcher.journal)
                for value in frame.get("account", pd.Series(dtype=object)).dropna() if value}
    account = next(iter(accounts)) if len(accounts) == 1 else ""
    company = identity[0] if len(identity) == 3 else ""
    bank_name = identity[1] if len(identity) == 3 else ""
    if not account and len(identity) == 3:
        account = identity[2]
    errors = reporter.error_collector.get_all_errors() if reporter.error_collector else []
    missing_amounts = [error for error in errors if error.get("type") in
                       {"金额解析失败", "日期解析失败", "空日期行", "方向解析失败"}]
    if missing_amounts:
        message += f"。汇总覆盖已读取的记录，另有 {len(missing_amounts)} 项原行信息未计入"
    if scope_conflicts:
        message = f"{scope_message}。已读到的收付分别列示，跨表差额留空。"
    overview = [
        ("公司", company), ("银行及账号", " · ".join(filter(None, (bank_name, account)))),
        ("银行流水", Path(bank_path).name if bank_path else ""), ("银行日记账", Path(journal_path).name if journal_path else ""),
        ("交易日期", f"{dates[0]:%Y-%m-%d} 至 {dates[-1]:%Y-%m-%d}" if dates else "无交易"),
        ("读入笔数", f"银行 {len(reporter.matcher.bank)} 笔，账上 {len(reporter.matcher.journal)} 笔"),
        ("月度核对结果", message), ("差异及范围说明" if automatic else "需要人工核对",
                                  len(pending) if automatic else sum(row[1] != "范围说明" for row in pending)),
        ("如何查看明细", "在差异说明点「查看全部…笔」。差异明细按编号列出原始记录。" if automatic
         else "在待核对点「查看全部…笔」。待核对明细按编号分组，浅蓝为范围说明，浅黄为核对事项。"),
        ("明显微小错报临界值", config.clearly_trivial_threshold), ("实际执行重要性水平", config.performance_materiality),
    ]
    if reporter.precheck_report:
        if reporter.matcher.balance_windows_checked and hasattr(reporter.matcher, "uncovered_months"):
            balance_result = "已使用可验证的余额区间核对交易，断点分别处理"
        elif reporter.matcher.balance_windows_checked:
            balance_result = "已核实本户及双方收支、期初期末余额，并逐日核对余额链"
        elif reporter.matcher.balance_scope_verified:
            balance_result = "已核实账户及总体余额。逐日余额链未能核实，未用于跨日或合笔确认"
        else:
            balance_result = "本次未同时核实账户、双方收支及连续余额，未启用余额配对"
        overview.append(("余额配对", balance_result))
    review_counts = getattr(reporter.matcher.llm_assistant, "review_counts", None)
    if review_counts is not None:
        overview.append(("Agent复核", "，".join(f"{key} {count} 项" for key, count in review_counts.items())))
    for error in errors:
        source = {"bank": "银行流水", "journal": "银行日记账"}.get(error.get("source_type"), "")
        reason = error.get("reason") or f"{error.get('type', '')}，{error.get('column', '')}原值 {error.get('original_value', '')}"
        overview.append(("原表读取说明", f"{source} 第{error.get('row', '')}行。{reason}"))
    pending_frame = pd.DataFrame(pending or [("", "", "本次无差异事项" if automatic else "本次无需人工核对", "", "", None, None, "", "", "")], columns=[
        "编号", "处理方式", "要查什么", "银行记录", "账上记录", "收入差额", "支出差额", "核查结论", "备注", "交易明细"])
    if automatic:
        pending_frame = pending_frame.drop(columns=["核查结论", "备注"]).rename(
            columns={"处理方式": "结果类型", "要查什么": "核对结果"})
    detail_columns = ["编号", "来源", "日期", "收入金额", "支出金额", "余额", "对方", "摘要",
                      "银行备注", "凭证号", "核对结果", "原表位置"]
    tables = {
        "核对总览": pd.DataFrame(overview, columns=["项目", "结果"]),
        "月度核对": _periods(monthly, "月份", "月末余额", monthly_notes),
        "日度核对": _periods(daily, "日期", "日末余额", daily_notes),
        report_tab: pending_frame,
        detail_tab: pd.DataFrame(review_details, columns=detail_columns),
        "交易明细": pd.DataFrame(details, columns=detail_columns),
    }
    plan = 新建簿计划({name: reporter._safe_table(table) for name, table in tables.items()})
    for sheet in plan.表:
        应用底层主题(sheet, 首列留白=False)
        sheet.冻结 = "A2"
        if sheet.名称 != "核对总览" and len(sheet.表):
            sheet.筛选 = f"A1:{字母(len(sheet.表.columns))}{len(sheet.表) + 1}"
    todo = plan.取表(report_tab)
    if pending:
        todo.表["交易明细"] = [
            "=" + 明细跳转公式(row, detail_tab, len(review_details) + 1, f"查看全部{count}笔")
            for row, (_, _, _, count) in enumerate(review_groups, 2)
        ]
    targeted = plan.取表(detail_tab)
    full = plan.取表("交易明细")
    for action, first_detail_row, first_review_row, count in review_groups:
        color = "DFEDFC" if action == "范围说明" else "FFEDBC"
        for sheet, start in ((targeted, first_review_row), (full, first_detail_row)):
            for index in range(count):
                for column in range(1, len(detail_columns) + 1):
                    sheet.盖格(start + index, column, 格样式(
                        填充色=color, 填充类型="solid", 加粗=True if column == 1 else None,
                        上框=("medium", "7891AD") if index == 0 else None,
                    ))
    for row in range(2, len(pending) + 2):
        if not automatic:
            for column in (8, 9):
                todo.盖格(row, column, 格样式(填充色="FFF2CC", 填充类型="solid"))
        todo.行计划取(row).高 = 68
    渲染簿计划(plan, output_path)
