"""相同摘要只能形成候选，不能单独证明审计核对关系已经闭合。"""
from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine


def _记录(金额, *, 对方="", 凭证=""):
    辅助文字 = {"摘要": "项目甲设备采购款"}
    if 对方:
        辅助文字["对方户名"] = 对方
    return {
        "date": pd.Timestamp("2026-08-31"),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": "项目甲设备采购款",
        "aux_text_fields": 辅助文字,
        "balance": None,
        "voucher_word": "记" if 凭证 else "",
        "voucher_no": 凭证,
        "original_idx": 1,
        "original_file_row": 2,
    }


def _核对(银行记录, 日记账记录):
    匹配器 = Matcher(
        pd.DataFrame(银行记录),
        pd.DataFrame(日记账记录),
        MatcherConfig(),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def test_唯一同日同额但只有相同具体摘要仍须列为疑点():
    匹配器 = _核对([_记录(-60000)], [_记录(-60000)])

    assert len(匹配器.selected_candidates) == 1
    候选 = 匹配器.selected_candidates[0]
    assert 候选.evidence["business_strength"] == 1
    assert 候选.processing_status is ProcessingStatus.FLAGGED
    assert "业务依据不足" in 候选.processing_reason


def test_明确对方加具体摘要的唯一同额关系仍可自动确认():
    匹配器 = _核对(
        [_记录(-300, 对方="甲公司")],
        [_记录(-300, 对方="甲公司")],
    )

    assert len(匹配器.selected_candidates) == 1
    assert (
        匹配器.selected_candidates[0].processing_status
        is ProcessingStatus.AUTO_CONFIRMED
    )


def test_完整凭证只有相同具体摘要时不能视为跨双方强证据():
    匹配器 = _核对(
        [_记录(-300)],
        [_记录(-100, 凭证="001"), _记录(-200, 凭证="001")],
    )

    候选 = next(
        项 for 项 in 匹配器.candidates
        if 项.bank_idxs == (0,)
        and 项.journal_idxs == (0, 1)
        and 项.evidence.get("atomic_voucher_group")
    )
    assert 候选.evidence.get("resolves_full_group") is False
    assert 候选.evidence.get("complete_business_group") is False
    assert 候选.processing_status is ProcessingStatus.FLAGGED
    assert "业务依据不足" in 候选.processing_reason
