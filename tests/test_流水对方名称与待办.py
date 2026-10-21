"""真实账上业务说明能证明交易对方时，不应制造人工待办。"""
from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine


def _记录(金额, 摘要, 辅助文字, 日期="2026-01-06", 原表行=5):
    return {
        "date": pd.Timestamp(日期),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 辅助文字,
        "balance": None,
        "voucher_word": "",
        "voucher_no": "",
        "original_idx": 原表行,
        "original_file_row": 原表行,
    }


def _核对(银行对方):
    bank = _记录(
        6344962.19, "货款", {"摘要": "货款", "对方户名": 银行对方}
    )
    journal_summary = "2026.1.6收广州交通集团物流有限公司货款"
    journal = _记录(
        6344962.19, journal_summary, {"业务说明": journal_summary}
    )
    matcher = Matcher(
        pd.DataFrame([bank]), pd.DataFrame([journal]), MatcherConfig(),
        logger=lambda _: None,
    )
    matcher.run()
    return matcher.selected_candidates[0]


def test_银行完整对方名称出现在账上业务说明时自动确认():
    candidate = _核对("广州交通集团物流有限公司")
    assert candidate.evidence["business_strength"] >= 2
    assert candidate.processing_status is ProcessingStatus.AUTO_CONFIRMED
    assert "广州交通集团物流有限公司" in candidate.evidence["business_basis"]


def test_短名称碰巧出现在业务说明时仍须人工核对():
    candidate = _核对("交通集团")
    assert candidate.evidence["business_strength"] < 2
    assert candidate.processing_status is ProcessingStatus.FLAGGED


def test_同单位同金额重复交易优先采用各自同日对应():
    party = "广州交通集团物流有限公司"
    dates = ("2026-01-06", "2026-01-17")
    banks = [_记录(200000, "货款", {"对方户名": party}, date, row)
             for date, row in zip(dates, (5, 15))]
    journals = [_记录(200000, f"收{party}货款", {"业务说明": f"收{party}货款"}, date, row)
                for date, row in zip(dates, (5, 15))]
    matcher = Matcher(pd.DataFrame(banks), pd.DataFrame(journals),
                      MatcherConfig(), logger=lambda _: None)
    matcher.run()
    assert len(matcher.selected_candidates) == 2
    assert all(candidate.date_span_days == 0 for candidate in matcher.selected_candidates)
    assert all(candidate.processing_status is ProcessingStatus.AUTO_CONFIRMED
               for candidate in matcher.selected_candidates)


def test_待办直接说明动作并标出原表行号():
    from 排版共享 import 核对动作, 记录描述

    bank = [{"日期": "2026-01-06", "摘要": "货款", "对方及凭证": "广州交通集团物流有限公司",
             "原始出处": "银行流水.xlsx · Sheet1 · 第5行"}]
    journal = [{"日期": "2026-01-06", "摘要": "收广州交通集团物流有限公司货款",
                "原始出处": "序时账.xlsx · Sheet1 · 第5行"}]
    action = 核对动作({"判断依据": "金额相同但依据不足"}, bank, journal, 0, 0)
    assert "对方" in action and "黄色格" in action
    ambiguous = 核对动作({"判断依据": "候选歧义"}, bank, journal, 0, 390.39)
    assert "其他对应方案" in ambiguous and "差额原因" in ambiguous
    assert "原表第5行" in 记录描述(bank)
