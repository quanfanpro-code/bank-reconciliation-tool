from itertools import chain

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from data_loader import DataLoader
from data_structures import MatcherConfig
from 列识别 import recognize_columns
from 运行核对 import SkillMatcher
from 银行核对入口 import AgentReview, run_accounts


def 出表(tmp_path, bank, journal):
    for name, records in (("现金发生", bank), ("财务记录", journal)):
        workbook = Workbook()
        workbook.active.title = "对账原件"
        for row in records:
            workbook.active.append(row)
        workbook.save(tmp_path / f"{name}.xlsx")
    result = run_accounts(
        tmp_path / "现金发生.xlsx", tmp_path / "财务记录.xlsx", tmp_path / "报告",
        MatcherConfig(), tmp_path / "会话", matcher_factory=SkillMatcher, log=lambda _: None,
    )
    assert not result["待处理"], result
    with open(result["完成"][0]["报告"], "rb") as stream:
        book = load_workbook(stream, read_only=True, data_only=True)
        try:
            return {sheet.title: list(sheet.values) for sheet in book}
        finally:
            book.close()


def 明细(values, key):
    headings, *rows = values
    return {str(row[headings.index(key)])[:10]: dict(zip(headings, row)) for row in rows}


def test_最后一笔之后没有收付且账上有完整期间仍有月底余额(tmp_path):
    bank = [["日期", "交易金额", "余额"], ["2030-05-17", -250, 1750]]
    journal = [["2030年3月至5月", "期初余额", "借方金额", "贷方金额", "期末余额"],
               [None, 2000, 0, 250, 1750], ["凭证日期", "借方金额", "贷方金额", "账面余额"],
               ["2030-05-17", None, 250, 1750]]
    report = 出表(tmp_path, bank, journal)
    months = 明细(report["月度核对"], "月份")
    assert list(months) == ["2030-03", "2030-04", "2030-05"]
    assert months["2030-03"]["账上月末余额"] == 2000
    assert months["2030-03"]["账上支出金额"] == 0
    assert months["2030-05"]["银行月末余额"] == months["2030-05"]["账上月末余额"] == 1750
    assert months["2030-05"]["余额差额"] == 0


def test_表明月中截止的流水仍只结转到已覆盖日期(tmp_path):
    bank = [["查询区间 2030-08-01 至 2030-08-14"], ["日期", "交易金额", "余额"], ["2030-08-14", 25, 950]]
    journal = [["2030年07月—08月", "期初账面余额", "借方发生额", "贷方发生额", "期末账面余额"],
               [None, 925, 25, 0, 950], ["凭证日期", "借方金额", "贷方金额", "账面余额"],
               ["2030-08-14", 25, None, 950]]
    report = 出表(tmp_path, bank, journal)
    august = 明细(report["月度核对"], "月份")["2030-08"]
    assert august["账上月末余额"] == 950 and august["银行月末余额"] is None
    assert "08月14日" in august["说明"]


def test_银行整月缺明细时账上照列且银行余额不能从上月直接抄入(tmp_path):
    bank = [["日期", "交易金额", "余额"], ["2030-03-20", -30, 770], ["2030-05-20", -20, 750]]
    journal = [["2030年3月至5月", "期初余额", "借方金额", "贷方金额", "期末余额"],
               [None, 800, 0, 50, 750], ["凭证日期", "借方金额", "贷方金额", "账面余额"],
               ["2030-03-20", None, 30, 770], ["2030-05-20", None, 20, 750]]
    report = 出表(tmp_path, bank, journal)
    april = 明细(report["月度核对"], "月份")["2030-04"]
    assert april["账上月末余额"] == 770 and april["银行月末余额"] is None
    assert april["账上支出金额"] == 0 and april["支出金额差额"] is None


def 核对(bank, journal, days):
    def load(side, values):
        raw = pd.DataFrame(values, columns=["日期", "交易金额", "摘要", "余额"])
        return DataLoader().standardize_data(raw, recognize_columns(raw, side).mapping, side)
    matcher = SkillMatcher(load("bank", bank), load("journal", journal),
                           MatcherConfig(tolerance_days=days, dfs_date_window=days), logger=lambda _: None)
    matcher.run()
    return matcher


@pytest.mark.parametrize("day1,day2,part1,part2", [("2031-02-15", "2031-03-09", 4, 9),
                                                                  ("2029-03-09", "2029-03-17", 150, 320)])
def test_同额跨日候选和同日合笔一起比较(day1, day2, part1, part2):
    bundle = part1 + part2
    bank = [(day1, -part1, "物业费", None), (day1, -part2, "物业费", None),
            (day2, -bundle, "物业费", None), (day2, -part1, "物业费", None)]
    journal = [(day1, -bundle, "物业费", None), (day2, -bundle-part1, "物业费", None)]
    matcher = 核对(bank, journal, 31)
    assert len(matcher.selected_candidates) == 2
    assert all(c.date_span_days == 0 and len(c.bank_idxs) == 2 for c in matcher.selected_candidates)
    assert sorted(chain.from_iterable(c.bank_idxs for c in matcher.selected_candidates)) == list(range(4))


def test_逐笔余额区分重复同额而且遵守日期设置():
    bank = [("2029-11-09", -17.23, "代扣", 320.77), ("2029-11-09", -17.23, "代扣", 303.54)]
    journal = [("2029-11-11", -17.23, "门禁卡费", 320.77), ("2029-11-11", -17.23, "门禁卡费", 303.54)]
    matcher = 核对(bank, journal, 2)
    assert len(matcher.selected_candidates) == 2
    for c in matcher.selected_candidates:
        b, = c.bank_idxs
        j, = c.journal_idxs
        assert matcher.bank.at[b, "balance"] == matcher.journal.at[j, "balance"]
    assert not 核对(bank, journal, 1).selected_candidates


@pytest.mark.parametrize("date_col,amount_col", [("收付发生于", "资金变化"), ("发生业务那天", "变动后的金额")])
def test_换了表头也把同日收付差额和原行写清(tmp_path, monkeypatch, date_col, amount_col):
    def answer(self, packet, **kwargs):
        if packet.get("kind") == "columns" and packet["side"] == "bank":
            return {"columns": {"bank": {"date": date_col, "amount": amount_col}}}, ""
        return None, "无补充判断"
    monkeypatch.setattr(AgentReview, "exchange", answer)
    bank = [[amount_col, "摘要", date_col], [200, "服务费", "2030-09-11"],
            [2705, "到账", "2030-09-27"], [-2705, "银行划款", "2030-09-27"]]
    journal = [["日期", "交易金额", "摘要"], ["2030-09-11", 200, "服务费"]]
    report = 出表(tmp_path, bank, journal)
    assert list(report) == ["核对总览", "月度核对", "日度核对", "差异说明", "差异明细", "交易明细"]
    headings, *rows = report["差异说明"]
    assert {"要查什么", "处理方式", "核查结论", "备注"}.isdisjoint(headings)
    assert sum(row[headings.index("收入差额")] or 0 for row in rows) == 2705
    assert sum(row[headings.index("支出差额")] or 0 for row in rows) == 2705
    assert "第3行" in 明细(report["日度核对"], "日期")["2030-09-27"]["说明"]
    assert "第4行" in 明细(report["日度核对"], "日期")["2030-09-27"]["说明"]
