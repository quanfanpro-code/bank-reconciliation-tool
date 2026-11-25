from decimal import Decimal
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

import pandas as pd

from data_structures import MatcherConfig, OverallControlResult, ProcessingStatus
from precision_engine import PrecisionEngine

SKILL_ENTRY = Path(__file__).resolve().parents[1] / ".agents/skills/bank-reconciliation-agent/scripts/运行核对.py"
spec = spec_from_file_location("reconciliation_skill", SKILL_ENTRY)
skill = module_from_spec(spec)
spec.loader.exec_module(skill)
review_spec = spec_from_file_location(
    "skill_review_items",
    SKILL_ENTRY.parent / "engine" / "复核事项.py",
)
review_module = module_from_spec(review_spec)
review_spec.loader.exec_module(review_module)
build_review_items = review_module.build_review_items

import sys
ENGINE_DIR = str(SKILL_ENTRY.parent / "engine")

def skill_module(name):
    sys.path.insert(0, ENGINE_DIR)
    sys.modules.pop(name, None)
    return __import__(name)



def frame(*rows):
    return pd.DataFrame([
        {
            "date": pd.Timestamp(day),
            "amount": Decimal(str(amount)),
            "amount_decimal": PrecisionEngine.to_integer_li(amount),
            "summary": summary,
            "aux_text_fields": {"摘要": summary, **fields},
            "original_idx": i,
            "original_file_row": i + 2,
            "voucher_no": "",
            "balance": fields.get("余额"),
        }
        for i, (day, amount, summary, fields) in enumerate(rows)
    ])


def run(bank, ledger, **kwargs):
    matcher = skill.SkillMatcher(bank, ledger, MatcherConfig(tolerance_days=3, dfs_date_window=3, performance_materiality=25000, clearly_trivial_threshold=2500), **kwargs)
    matcher.run()
    return matcher


def test_逐笔容差内唯一的弱摘要大额由程序确认():
    bank = frame(("2026-01-31", -520000.70, "转账", {}))
    ledger = frame(("2026-02-03", -520000.70, "天然气款", {}))
    control = OverallControlResult(scope_limited=True, reasons=("核对币种未完全验证",))
    matcher = run(bank, ledger, overall_control=control)
    assert [(c.bank_idxs, c.journal_idxs, c.processing_status) for c in matcher.selected_candidates] == [
        ((0,), (0,), ProcessingStatus.AUTO_CONFIRMED)
    ]
    assert build_review_items(matcher)[0]["核对方式"] == "自动确认"


def test_超过日期容差或方向相反不逐笔对上():
    bank = frame(("2026-01-01", -520000.70, "转账", {}))
    for row in (("2026-01-05", -520000.70, "转账", {}), ("2026-01-01", 520000.70, "转账", {})):
        assert not run(bank, frame(row)).selected_candidates


def test_递归剪枝查到的唯一合笔由程序确认():
    bank = frame(("2026-01-20", -150, "转账", {}), ("2026-01-21", -147, "转账", {}))
    ledger = frame(("2026-01-22", -297, "付款", {}))
    matcher = run(bank, ledger)
    assert [(c.bank_idxs, c.journal_idxs) for c in matcher.selected_candidates] == [((0, 1), (0,))]
    assert build_review_items(matcher)[0]["核对方式"] == "自动确认"


def test_等额重复整组闭合但不捏造逐笔关系():
    bank = frame(("2026-01-01", 40000, "转账", {}), ("2026-01-02", 40000, "转账", {}))
    ledger = frame(("2026-01-02", 40000, "收款", {}), ("2026-01-03", 40000, "收款", {}))
    matcher = run(bank, ledger)
    assert [(c.bank_idxs, c.journal_idxs, c.processing_status) for c in matcher.selected_candidates] == [
        ((0, 1), (0, 1), ProcessingStatus.AUTO_CONFIRMED)
    ]


def test_未超过重要性的单边差额不进差异说明():
    bank = frame(("2026-01-08", -1000, "手续费", {}), ("2026-01-09", -10000, "设备款", {}))
    ledger = frame(("2026-01-08", -1, "无关", {}))
    modes = {item["影响金额"]: item["核对方式"] for item in build_review_items(run(bank, ledger))}
    assert modes[Decimal("1000")] == "留存备查"
    assert modes[Decimal("10000")] == "留存备查"
    assert "差异说明" not in modes.values()


def test_刚好等于重要性不算超过():
    bank = frame(("2026-01-08", -25000, "设备款", {}))
    ledger = frame(("2026-01-08", -1, "无关", {}))
    items = build_review_items(run(bank, ledger))
    assert items[0]["影响金额"] == Decimal("25000")
    assert items[0]["核对方式"] == "留存备查"


def test_超过重要性或同月累计超过才进差异说明():
    bank = frame(("2026-01-08", -30000, "设备款", {}))
    ledger = frame(("2026-01-08", -1, "无关", {}))
    assert build_review_items(run(bank, ledger))[0]["核对方式"] == "差异说明"
    pieces = frame(
        ("2026-02-01", -10000, "材料", {}),
        ("2026-02-02", -10000, "运费", {}),
        ("2026-02-03", -10000, "服务费", {}),
    )
    other = frame(("2026-02-01", -1, "无关", {}))
    flagged = [item for item in build_review_items(run(pieces, other)) if item["核对方式"] == "差异说明"]
    assert flagged
    assert sum(item["影响金额"] for item in flagged) == Decimal("30000")


def test_手续费净额已说明时不进待办且月度不留未解释差额(tmp_path):
    from openpyxl import load_workbook
    Reporter = skill_module("核对报告") and skill_module("reporter").Reporter

    bank = frame(("2026-01-10", 970, "货款净额", {"对方户名": "甲公司", "业务编号": "P01"}))
    ledger = frame(
        ("2026-01-10", 1000, "货款", {"对方户名": "甲公司", "业务编号": "P01"}),
        ("2026-01-10", -30, "手续费", {"对方户名": "甲公司", "业务编号": "P01"}),
    )
    matcher = run(bank, ledger)
    assert all(item["核对方式"] != "差异说明" for item in build_review_items(matcher))
    target = tmp_path / "报告.xlsx"
    Reporter(matcher)._generate_report(str(target))
    book = load_workbook(target, data_only=True)
    headers = [cell.value for cell in book["月度核对"][1]]
    row = [cell.value for cell in book["月度核对"][2]]
    record = dict(zip(headers, row))
    book.close()
    assert Decimal(str(record["收入金额差额"])) == 0
    assert Decimal(str(record["支出金额差额"])) == 0
    assert "手续费净额已说明" in str(record["说明"])


def test_同额冲突和差额未被掩盖():
    bank = frame(("2026-01-01", 40000, "转账", {"对方户名": "甲有限公司"}),
                 ("2026-01-01", -80000, "付款", {}))
    ledger = frame(("2026-01-01", 40000, "收款", {"对方户名": "乙有限公司"}),
                   ("2026-01-01", -50000, "转账", {}))
    items = build_review_items(run(bank, ledger))
    assert any(item["核对方式"] == "差异说明" and "对方户名" in item["判断依据"] for item in items)
    assert sum(item["银行支出"] - item["日记账支出"] for item in items) == Decimal("30000")


def test_某月一侧无明细只集中说明一次():
    bank = frame(("2026-01-31", 50000, "转账", {}), ("2026-03-01", 10000, "转账", {}))
    ledger = frame(("2026-01-31", 50000, "收款", {}),
                   ("2026-02-01", 30000, "收款", {}), ("2026-02-02", -40000, "付款", {}),
                   ("2026-03-01", 10000, "收款", {}))
    items = build_review_items(run(bank, ledger))
    assert [(i["银行索引"], i["日记账索引"]) for i in items if i["核对方式"] == "范围说明"] == [((), (1, 2))]
    assert sum(i["核对方式"] == "自动确认" for i in items) == 2

def test_语义冲突由Agent读原文决定():
    class Review:
        def __init__(self, status, quote):
            self.status, self.quote = status, quote
            self.review_counts = {"已完成": 0, "未复核": 0}

        def exchange(self, packet):
            assert packet["kind"] == "purpose"
            return {"verdicts": {
                choice["candidate_id"]: {
                    "status": self.status,
                    "bank": {"field": "摘要", "quote": self.quote},
                    "journal": {"field": "摘要", "quote": "购设备款"},
                } for choice in packet["candidates"]
            }}, ""

    bank = frame(("2026-01-20", -50000, "发放工资", {}))
    ledger = frame(("2026-01-20", -50000, "购设备款", {}))
    responder = Review("conflict", "发放工资")
    result = build_review_items(run(bank, ledger, llm_assistant=responder))
    assert result[0]["核对方式"] == "差异说明"
    assert "发放工资" in result[0]["判断依据"] and "购设备款" in result[0]["判断依据"]
    invalid = Review("conflict", "银行原表没有这句话")
    result = build_review_items(run(bank, ledger, llm_assistant=invalid))
    assert result[0]["核对方式"] == "自动确认"
    assert invalid.review_counts["未复核"]


def test_银行同日等额收付仍保留两个方向的实际差额():
    bank = frame(("2026-01-20", 50000, "转账", {"对方户名": "甲有限公司", "余额": Decimal("51000")}),
                 ("2026-01-20", -50000, "转账", {"对方户名": "甲有限公司", "余额": Decimal("1000")}),
                 ("2026-01-21", 1000, "收款", {"余额": Decimal("2000")}))
    ledger = frame(("2026-01-21", 1000, "收款", {"余额": Decimal("2000")}))
    items = build_review_items(run(bank, ledger))
    actions = [i for i in items if i["核对方式"] == "差异说明"]
    assert len(actions) == 1
    assert actions[0]["银行索引"] == (0, 1)
    assert actions[0]["银行收入"] == actions[0]["银行支出"] == Decimal("50000")

def test_双方整月无行也需要查询范围():
    bank = frame(("2025-05-31", 12, "收款", {}), ("2025-07-01", 25, "收款", {}))
    journal = frame(("2025-05-31", 12, "收款", {}), ("2025-07-01", 25, "收款", {}))
    matcher = run(bank, journal)
    assert matcher.uncovered_months == {"bank": {pd.Period("2025-06")}, "journal": {pd.Period("2025-06")}}

def test_范围说明的全部明细按编号着色并可直接跳转(tmp_path):
    from openpyxl import load_workbook
    _验证五表结构 = skill_module("Excel验证")._验证五表结构
    Reporter = skill_module("reporter").Reporter

    bank = frame(("2025-05-31", 2800, "收款", {}), ("2025-07-01", 3200, "收款", {}))
    journal = frame(("2025-05-31", 2800, "收款", {}),
                    *(("2025-06-14", i, "汇款", {}) for i in (13, 24, 37, 42, 56)),
                    ("2025-07-01", 3200, "收款", {}))
    matcher = run(bank, journal)
    target = tmp_path / "报告.xlsx"
    Reporter(matcher)._generate_report(str(target))
    assert _验证五表结构(target)["工作表数"] == 6
    book = load_workbook(target)
    review, subset, full = (book[name] for name in ("差异说明", "差异明细", "交易明细"))
    number = review["A2"].value
    assert review["C2"].value == "2025年6月银行缺这些日期的流水，账上有5笔。收付差额无法计算。"
    assert f"第5–6行（编号{number}）" in review["E2"].value
    assert "差异明细" in review["H2"].value and "查看全部5笔" in review["H2"].value
    assert subset.max_row == 6
    assert {subset.cell(row, 1).value for row in range(2, 7)} == {number}
    assert {cell.fill.fgColor.rgb[-6:] for row in subset.iter_rows(min_row=2, max_row=6) for cell in row} == {"DFEDFC"}
    assert [row for row in full.iter_rows(min_row=2, values_only=True) if row[0] == number] == list(subset.values)[1:]
    book.close()
