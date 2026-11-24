import pandas as pd
import pytest

from 运行核对 import SkillMatcher
from data_loader import DataLoader
from data_structures import MatcherConfig
from 列识别 import recognize_columns
from 复核事项 import build_review_items


def 核对(流水, 账上, *, 天数=3, 答复=None, 凭据=None):
    def 标准化(内容, 来源):
        raw = pd.DataFrame(内容, columns=["日期", "交易金额", "摘要", "对方户名", "订单编号"])
        return DataLoader().standardize_data(raw, recognize_columns(raw, 来源).mapping, 来源)
    matcher = SkillMatcher(标准化(流水, "bank"), 标准化(账上, "journal"),
                           MatcherConfig(tolerance_days=天数, dfs_date_window=天数),
                           llm_assistant=答复, logger=lambda _: None)
    matcher.coverage_evidence = 凭据 or {}
    matcher.run()
    return matcher


@pytest.mark.parametrize("days,matched", [(31, True), (3, False), (0, False)])
def test_月度合笔不能绕过本次配置的日期范围(days, matched):
    bank = [("2025-05-01", -6000, "工资", "", ""), ("2025-05-01", -4000, "工资", "", "")]
    ledger = [("2025-05-30", -10000, "工资", "", "")]
    matcher = 核对(bank, ledger, 天数=days)
    assert bool(matcher.selected_candidates) is matched


def test_等额退款也要核对原订单():
    anchors = [("2025-05-01", 100, "收款", "", ""), ("2025-05-30", 50, "收款", "", "")]
    bank = [*anchors, ("2025-05-03", -100000, "支付货款", "甲公司", "A20250503"),
            ("2025-05-03", 100000, "货款冲正", "甲公司", "B20250503")]
    matcher = 核对(bank, anchors)
    assert not any(item["程序结论"] == "银行冲正已闭合" for item in build_review_items(matcher))


def test_本方账号相同也要核对交易对方账号():
    def a(账号):
        raw = pd.DataFrame([["2025-05-05", 100, "001122", 账号, "货款"]],
                           columns=["日期", "交易金额", "本方账号", "对方账号", "摘要"])
        return DataLoader().standardize_data(raw, recognize_columns(raw, "bank").mapping, "bank")
    matcher = SkillMatcher(a("002211"), a("002299"), MatcherConfig(), logger=lambda _: None)
    matcher.run()
    assert not matcher.selected_candidates


def test_明细缺口内余额不能盖过两表相同订单():
    day = pd.Timestamp("2025-05-05")
    def read(raw, side):
        frame = pd.DataFrame(raw, columns=["日期", "交易金额", "摘要", "订单编号"])
        return DataLoader().standardize_data(frame, recognize_columns(frame, side).mapping, side)
    bank = read([(day, 100, "货款", "ORD1001")], "bank")
    journal = read([(day, 100, "收款", "ORD1002"), (day, 100, "货款", "ORD1001")], "journal")
    bad = int(journal.index[journal["summary"].eq("收款")][0])
    matcher = SkillMatcher(bank, journal, MatcherConfig(), logger=lambda _: None,
                           verified_balance_pairs={(0, bad): "试算余额相同"})
    matcher.coverage_evidence = {"bank": {"periods": [(day, day, "不完整", "银行原表")]}}
    matcher.run()
    selected, = matcher.selected_candidates
    assert not selected.evidence.get("verified_balance_match")
    assert journal.loc[list(selected.journal_idxs), "summary"].tolist() == ["货款"]


class 原文回复:
    def __init__(self, second_timeout=False):
        from collections import Counter
        self.review_counts = Counter()
        self.calls = 0
        self.second_timeout = second_timeout

    def exchange(self, packet):
        self.calls += 1
        if self.second_timeout and self.calls > 1:
            return None, "超时"
        return {"verdicts": {c["candidate_id"]: {"status": "conflict",
                "bank": {"field": "摘要", "quote": "支付1月工资"},
                "journal": {"field": "摘要", "quote": "支付2月工资"}} for c in packet["candidates"]}}, ""


def test_后续超时不能抹掉原文证实的不同月份():
    bank = [("2025-05-01", -10000, "支付1月工资", "", "")]
    ledger = [("2025-05-01", -10000, "支付2月工资", "", "")]
    matcher = 核对(bank, ledger, 答复=原文回复(second_timeout=True))
    assert not any(c.processing_status.value == "自动确认" for c in matcher.selected_candidates)


def test_加入不可能参与的交易不改变整组选择():
    from collections import Counter
    from data_structures import LLMDecisionRecord
    class 暂无文字证据:
        review_counts = Counter()
        def evaluate_candidates(self, request):
            return LLMDecisionRecord(request.request_id, fallback_used=True)
    day = "2025-05-11"
    bank = [(day, -10000, "付款", "", "")]
    ledger = [(day, -amount, "付款", "", "") for amount in (7000, 3000, 6000, 4000)]
    def 采用(账):
        mt = 核对(bank, 账, 答复=暂无文字证据())
        c, = mt.selected_candidates
        return sorted(mt.journal.loc[list(c.journal_idxs), "amount"])
    extras = [(day, -amount, "付款", "", "") for amount in range(7501, 7528)]
    assert 采用(ledger) == 采用(extras + ledger)


def test_未展示的候选也核对原表后自动选中一月款(tmp_path):
    from 银行核对入口 import AgentReview
    class 正常答复(AgentReview):
        def exchange(self, packet, **kwargs):
            selected, = [c for c in packet["candidates"] if c["journal_fields"]["摘要"] == "2026年1月天然气款"]
            return {"selected_candidate_id": selected["candidate_id"], "semantic_score": 100, "reason": "原文相同",
                    "source_evidence": {"bank_field": "摘要", "journal_field": "摘要", "quote": "2026年1月天然气款"}}, ""
    日期 = "2026-02-05"
    内容 = "2026年1月天然气款"
    银行 = [(日期, 150000, 内容, "", "")]
    账上 = [*银行, *((日期, 150000, f"2025年{month}月天然气款", "", "") for month in range(8, 13))]
    matcher = 核对(银行, 账上, 答复=正常答复(tmp_path))
    candidate, = matcher.selected_candidates
    assert candidate.evidence.get("agent_confirmed")
    assert matcher.journal.loc[list(candidate.journal_idxs), "summary"].tolist() == [内容]
