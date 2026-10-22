"""把已有匹配记录整理为业务结论、待答问题和一处可读的完整证据。"""
from collections import Counter, defaultdict
from decimal import Decimal

import pandas as pd


def 收支(records):
    amounts = [Decimal(str(r["金额"])) for r in records]
    return (sum((max(a, Decimal(0)) for a in amounts), Decimal(0)),
            sum((max(-a, Decimal(0)) for a in amounts), Decimal(0)))


def 记录内容(record):
    return "；".join(str(v) for v in (
        pd.Timestamp(record["日期"]).strftime("%Y-%m-%d"), record.get("摘要"),
        record.get("对方及凭证"), f"金额 {Decimal(str(record['金额'])):,.2f}元",
    ) if v is not None and str(v).strip())


def 组成说明(records):
    if not records:
        return "本侧尚未找到对应记录"
    income, expense = 收支(records)
    detail = "\n".join(记录内容(r) for r in records[:3])
    if len(records) > 3:
        detail += f"\n另{len(records)-3}笔，点击完整证据直接查看全部组成"
    return f"{len(records)}笔；收入合计 {income:,.2f}；支出合计 {expense:,.2f}\n{detail}"


def 归纳事项(item, bank, journal):
    bi, be = 收支(bank)
    ji, je = 收支(journal)
    di, de = bi - ji, be - je
    parts = []
    if item["程序结论"] == "银行冲正已闭合":
        return (f"同日同对方唯一付款{be:,.2f}元与银行明确冲正退回{bi:,.2f}元闭合；净支付为零。"
                "银行原始收入与支出分别保留，月度及每日差额不抵销。\n" + str(item["判断依据"]),
                "本项无需人工寻找另一笔账上支付，失败支付及银行退回的完整记录留存备查。")
    if item["程序结论"] == "账内更正已闭合":
        return ("明确引用原凭证的冲销重记已闭合；原付款维持原对应。本次账内收入与支出分别为"
                f"{bi or ji:,.2f}元、{be or je:,.2f}元，净额为零；月度及每日原始收支差额保留。\n" + str(item["判断依据"]),
                "本项无需另找银行交易或人工填写；冲销、重记及引用原凭证依据已列示，留存备查。")
    if not bank or not journal:
        side = "银行" if bank else "账上"
        parts.append(f"真正单边：仅{side}有{len(bank or journal)}笔；收入{bi or ji:,.2f}元、支出{be or je:,.2f}元")
    else:
        parts.append(f"收入：银行{bi:,.2f}－账上{ji:,.2f}＝{di:,.2f}元；支出：银行{be:,.2f}－账上{je:,.2f}＝{de:,.2f}元")
        if len(bank) != len(journal):
            parts.append(f"银行{len(bank)}笔、账上{len(journal)}笔；属于拆分/汇总关系，笔数不同本身不证明漏记")
        bank_days = {pd.Timestamp(r["日期"]).strftime("%Y-%m-%d") for r in bank}
        journal_days = {pd.Timestamp(r["日期"]).strftime("%Y-%m-%d") for r in journal}
        if bank_days != journal_days:
            parts.append("双方记账日期不同：银行" + "、".join(sorted(bank_days)) + "；账上" + "、".join(sorted(journal_days)))
        if di or de:
            # 只归纳金额组成，不能用同额抵销冒充已建立逐笔关系。
            b = Counter(Decimal(str(r["金额"])) for r in bank)
            j = Counter(Decimal(str(r["金额"])) for r in journal)
            for label, remaining in (("银行", b-j), ("账上", j-b)):
                if remaining:
                    values = [f"{amount:,.2f}元×{count}笔" for amount, count in sorted(remaining.items())]
                    parts.append(label + "金额比较剩余组成：" + "、".join(values[:6]) + ("等，完整组成见证据" if len(values) > 6 else ""))
                    source = bank if label == "银行" else journal
                    left = remaining.copy()
                    clues = []
                    for record in source:
                        amount = Decimal(str(record["金额"]))
                        if left[amount] > 0:
                            left[amount] -= 1
                            clues.append(记录内容(record))
                    parts.append(label + "差额定位记录：" + "；".join(clues[:3]) + ("；其余见完整证据" if len(clues)>3 else ""))
            parts.append("上述金额比较仅帮助定位差额，不代替业务对应判断")
    reason = str(item.get("判断依据") or "")
    if "冲突" in reason:
        parts.append("业务依据存在明确冲突：" + reason)
    elif item.get("核对方式", "").startswith("人工") and reason:
        parts.append("尚未确认原因：" + reason)
    if not bank or not journal:
        question = "这笔单边收付是否在另一侧漏记、跨日/跨月记账或不属于本账户？请查原始流水、序时账及凭证，在黄色格记录核查结论和原因；尚未查明可留空。"
    elif di or de:
        question = f"这些记录是否属于同一业务？收入差{di:,.2f}元、支出差{de:,.2f}元各是什么原因，是否有组成漏列？请查原始凭证，自由记录核查结论及差额原因；填写不会改变原始差额。"
    elif "冲突" in reason:
        question = "双方列示的对方、用途或业务编号为何不同，能否证明属于同一业务？请查原始凭证，在黄色格记录核查结论，具体依据填备注。"
    elif (any("工资" in str(r.get("摘要", "")) for r in journal)
          and ("业务依据不足" in reason or "缺少足以证明" in reason)):
        parts.append("账上列为工资，银行记录未提供可核实的工资用途或与账上共同的员工标识；金额组成相同尚不能证明人员归属")
        question = "请用工资表确认本处列出的银行收款人名单，是否就是账上完整凭证所记的工资及感恩父母基金；确认整组范围后在黄色格填写，差额或名单变化写备注。"
    elif "歧义" in reason:
        question = "本项存在多种可能对应，程序列出的线索并未穷尽。请查双方完整记录及原始凭证，自由记录查明的情况和依据，无需从候选中选择。"
    else:
        question = "双方金额相同，日期、对方和用途是否足以证明同一业务？请查原始记录，在黄色格自由记录核查结论，具体依据填备注。"
    if item.get("核对方式") == "自动确认":
        question = "程序已有充分证据确认，本项无需人工填写；完整组成留存供查阅。"
    elif item.get("核对方式") == "留存备查":
        question = "本项按既定重要性和抽样规则留存备查，本次无需人工填写；此处不代表已经确认对应。"
    return "\n".join(parts), question


def 整理业务证据(tables):
    records = defaultdict(lambda: {"银行流水": [], "序时账": []})
    for r in tables["核对明细"].to_dict("records"):
        records[r["事项编号"]][r["来源"]].append(r)
    alternatives = defaultdict(list)
    for r in tables["其他对应线索"].to_dict("records"):
        alternatives[r["事项编号"]].append(r)
    items = tables["复核事项索引"]
    evidence = []
    explanations = {}
    for item in items.to_dict("records"):
        identity = item["事项编号"]
        bank, journal = records[identity]["银行流水"], records[identity]["序时账"]
        summary, question = 归纳事项(item, bank, journal)
        explanations[identity] = (summary, question, 组成说明(bank), 组成说明(journal))
        方案数 = len({r['线索编号'] for r in alternatives[identity]})
        evidence.append({"核对编号": item["核对编号"], "事项编号": identity, "行别": "事项结论与合计",
                         "结论与待答问题": item["程序结论"] + "；" + item["核对方式"] + "\n" + summary + "\n" + question,
                         "银行记录": 组成说明(bank), "序时账记录": 组成说明(journal),
                         "证据说明": item["判断依据"],
                         "其他线索": f'{方案数}套参考线索，点击查看（未穷尽所有可能，无需选择）' if 方案数 else ''})
        for side, group in (("银行", bank), ("序时账", journal)):
            for r in group:
                amount = Decimal(str(r["金额"]))
                evidence.append({"核对编号": item["核对编号"], "事项编号": identity, "行别": side + "组成",
                                 side + "记录": 记录内容(r), side + "收入": float(max(amount, Decimal(0))),
                                 side + "支出": float(max(-amount, Decimal(0))), "证据说明": r["原始出处"]})
    for pos, name in enumerate(("程序归纳", "具体问题", "双方银行说明", "双方账上说明")):
        items[name] = items["事项编号"].map(lambda identity: explanations[identity][pos])
    tables["事项证据"] = pd.DataFrame(evidence, columns=["核对编号", "行别", "结论与待答问题", "银行记录", "银行收入", "银行支出", "序时账记录", "序时账收入", "序时账支出", "证据说明", "返回填写", "事项编号", "其他线索"])

    daily_records = defaultdict(list)
    for sides in records.values():
        for group in sides.values():
            for r in group:
                daily_records[pd.Timestamp(r["日期"]).strftime("%Y-%m-%d")].append(r)
    days = []
    for row in tables["每日统计"].to_dict("records"):
        day = pd.Timestamp(row["日期"]).strftime("%Y-%m-%d")
        group = daily_records[day]
        reasons = []
        for name, label in (("收入金额差额", "收入差"), ("支出金额差额", "支出差"), ("余额差额", "余额差")):
            value = row.get(name)
            if value is not None and pd.notna(value) and abs(Decimal(str(value))) >= Decimal("0.0001"):
                reasons.append(f"{label}{Decimal(str(value)):,.2f}元")
        for name, label in (("收入笔数差额", "收款笔数差"), ("支出笔数差额", "付款笔数差")):
            value = row.get(name)
            if value is not None and pd.notna(value) and value:
                reasons.append(f"{label}{int(value)}笔（可能拆分或跨日）")
        explanation = "；".join(reasons) or "当日收支及已提供余额无差额；并不单独证明逐笔对应"
        days.append({"日期": day, "行别": "当日结论", "业务内容": explanation,
                     "相关事项与核查说明": "下列为该日双方全部交易；跨日、拆分或单边原因由相关事项完整证据说明。"})
        for r in sorted(group, key=lambda r: (r["来源"] != "银行流水", r["核对编号"])):
            summary, question, _, _ = explanations[r["事项编号"]]
            amount = Decimal(str(r["金额"]))
            days.append({"日期": day, "行别": r["来源"], "核对编号": r["核对编号"],
                         "业务内容": 记录内容(r), "收入": float(max(amount, Decimal(0))),
                         "支出": float(max(-amount, Decimal(0))), "相关事项与核查说明": summary,
                         "事项编号": r["事项编号"]})
    tables["每日差异证据"] = pd.DataFrame(days, columns=["日期", "行别", "核对编号", "业务内容", "收入", "支出", "相关事项与核查说明", "查看完整事项", "事项编号"])
    return tables
