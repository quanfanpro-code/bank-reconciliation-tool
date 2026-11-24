from types import SimpleNamespace

import pandas as pd

from reporter import Reporter
from data_structures import MatcherConfig
from 运行核对 import SkillMatcher
from 银行核对入口 import run_accounts
from openpyxl import load_workbook
from tests.test_Skill全面修复输入 import 原表, 表头


def test_一侧没提供余额仍保留银行原值():
    days = pd.to_datetime(["2025-05-12", "2025-05-13"])
    bank = pd.DataFrame({"date": days, "amount": [100, 50], "balance": [1100, 1150]})
    ledger = pd.DataFrame({"date": days, "amount": [100, 50]})
    matcher = SimpleNamespace()
    daily = Reporter.calculate_daily_stats(SimpleNamespace(matcher=matcher), bank, ledger)
    assert list(daily["balance_bank"]) == [1100, 1150]
    assert daily["balance_journal"].isna().all()


def 跑完整核对(tmp_path, bank_rows, journal_rows):
    bank, ledger = tmp_path / "流水.xlsx", tmp_path / "日记账.xlsx"
    原表(bank, {"核对原页": bank_rows})
    原表(ledger, {"核对原页": journal_rows})
    result = run_accounts(bank, ledger, tmp_path / "报告",
                          MatcherConfig(performance_materiality=500, clearly_trivial_threshold=20),
                          tmp_path / "会话", matcher_factory=SkillMatcher, log=lambda _: None)
    assert not result["待处理"], result
    with open(result["完成"][0]["报告"], "rb") as file:
        book = load_workbook(file, read_only=True, data_only=True)
        try:
            return {name: list(book[name].values) for name in book.sheetnames}
        finally:
            book.close()


def test_部分月份与同额余额断差写入原报告(tmp_path):
    rows = [["查询起止：2025-05-12 至 2025-05-13"], [*表头, "余额"],
            ["2025-05-12", 100, "001122", "CNY", "收款", 1100],
            ["2025-05-13", 50, "001122", "CNY", "收款", 4100]]
    report = 跑完整核对(tmp_path, rows, rows)
    monthly, daily = report["月度核对"], report["日度核对"]
    assert monthly[1][monthly[0].index("银行月末余额")] is None
    assert "05月13日" in " ".join(str(value) for value in monthly[1])
    assert "2,950.00" in " ".join(str(value) for row in daily[1:] for value in row)


def test_两个币种不同保留各自原表且不计算跨币差额(tmp_path):
    report = 跑完整核对(tmp_path, [表头, ["2025-05-31", "USD 100", "001122", "USD", "销售款"]],
                        [表头, ["2025-05-31", "CNY 100", "001122", "CNY", "销售款"]])
    monthly = report["月度核对"]
    assert monthly[1][monthly[0].index("收入金额差额")] is None
    assert "币种" in " ".join(str(value) for row in report["核对总览"] for value in row)


def test_异月大额差异不把手续费转成人工(tmp_path):
    anchors = [["2025-04-30", 10, "001122", "CNY", "回款"], ["2025-05-31", 30, "001122", "CNY", "回款"]]
    bank = [表头, *anchors, ["2025-04-18", 5000, "001122", "CNY", "银行代收"],
            ["2025-05-15", -200, "001122", "CNY", "手续费"]]
    report = 跑完整核对(tmp_path, bank, [表头, *anchors])
    pending = report["待核对"]
    assert all(row[1] == "差异说明" for row in pending[1:])
    assert any(row[pending[0].index("支出差额")] == 200 for row in pending[1:])


def test_只有一侧可计算金额仍在原月表照列(tmp_path):
    report = 跑完整核对(tmp_path, [表头, ["2025-05-31", "未列金额", "001122", "CNY", "销售款"]],
                        [表头, ["2025-05-31", 100, "001122", "CNY", "销售款"]])
    monthly = report["月度核对"]
    assert monthly[1][monthly[0].index("账上收入金额")] == 100
    assert monthly[1][monthly[0].index("银行收入金额")] is None
    assert monthly[1][monthly[0].index("收入金额差额")] is None
    assert "第2行" in " ".join(str(value) for row in report["核对总览"] for value in row)


def test_多笔相同金额仍能汇总收付(tmp_path):
    entries = [[f"2025-05-{day:02d}", 100, "001122", "CNY", "收款"] for day in (5, 12, 26)]
    report = 跑完整核对(tmp_path, [表头, *entries], [表头, *entries])
    header, month = report["月度核对"]
    assert month[header.index("银行收入金额")] == 300
    assert month[header.index("收入金额差额")] == 0


def test_同一天原表出现两个余额时把两笔原值写进日表(tmp_path):
    header = [*表头, "余额"]
    date, amount = "2025-05-31", 50
    prefixes = [["查询起止：2025-05-01 至 2025-05-31"], ["2025-05-31 余额 1150"], ["2025-05-31 余额 1100"]]
    row = [date, amount, "001122", "CNY", "收款", 1150]
    report = 跑完整核对(tmp_path, [*prefixes, header, row], [header, row])
    text = " ".join(str(value) for row in report["日度核对"] for value in row)
    assert "1,150.00" in text and "1,100.00" in text and "50.00" in text


def test_收付相同仍在明细保留重复流水号线索(tmp_path):
    header = [*表头, "交易流水号"]
    a = ["2025-05-31", -1000, "001122", "CNY", "付款", "TX0101"]
    report = 跑完整核对(tmp_path, [header, a, a], [header, a, a])
    details = report["交易明细"]
    assert len(details) == 5
    assert sum("流水号" in str(value) for row in details[1:] for value in row) >= 2
