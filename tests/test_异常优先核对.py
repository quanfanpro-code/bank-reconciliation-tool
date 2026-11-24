from decimal import Decimal
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path

import pandas as pd

from data_structures import MatcherConfig, OverallControlResult, ProcessingStatus
from precision_engine import PrecisionEngine
from 复核事项 import build_review_items

SKILL_ENTRY = Path(__file__).resolve().parents[1] / ".agents/skills/bank-reconciliation-agent/scripts/运行核对.py"
spec = spec_from_file_location("reconciliation_skill", SKILL_ENTRY)
skill = module_from_spec(spec)
spec.loader.exec_module(skill)


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


def test_同额冲突和差额未被掩盖():
    bank = frame(("2026-01-01", 40000, "转账", {"对方户名": "甲有限公司"}),
                 ("2026-01-01", -80000, "付款", {}))
    ledger = frame(("2026-01-01", 40000, "收款", {"对方户名": "乙有限公司"}),
                   ("2026-01-01", -50000, "转账", {}))
    items = build_review_items(run(bank, ledger))
    assert any(item["核对方式"] == "人工全查" and "对方户名" in item["判断依据"] for item in items)
    assert sum(item["银行支出"] - item["日记账支出"] for item in items) == Decimal("30000")


def test_某月一侧无明细只集中说明一次():
    bank = frame(("2026-01-31", 50000, "转账", {}), ("2026-03-01", 10000, "转账", {}))
    ledger = frame(("2026-01-31", 50000, "收款", {}),
                   ("2026-02-01", 30000, "收款", {}), ("2026-02-02", -40000, "付款", {}),
                   ("2026-03-01", 10000, "收款", {}))
    items = build_review_items(run(bank, ledger))
    assert [(i["银行索引"], i["日记账索引"]) for i in items if i["核对方式"] == "人工全查"] == [((), (1, 2))]
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
    assert result[0]["核对方式"] == "人工全查"
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
    actions = [i for i in items if i["核对方式"] == "人工全查"]
    assert len(actions) == 1
    assert actions[0]["银行索引"] == (0, 1)
    assert actions[0]["银行收入"] == actions[0]["银行支出"] == Decimal("50000")
