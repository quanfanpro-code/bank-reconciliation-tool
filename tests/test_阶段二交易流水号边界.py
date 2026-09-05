from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine


def _记录(金额, 流水号):
    return {
        "date": pd.Timestamp("2026-08-10"),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": "项目回款",
        "aux_text_fields": {"摘要": "项目回款", "交易流水号": 流水号},
        "voucher_word": "",
        "voucher_no": "",
    }


def _表格(记录集):
    return pd.DataFrame(
        [
            dict(行, original_idx=序号 + 1, original_file_row=序号 + 2)
            for 序号, 行 in enumerate(记录集)
        ]
    )


def _核对(日记账流水号):
    匹配器 = Matcher(
        _表格([_记录(100, "R1")]),
        _表格([_记录(40, "R1"), _记录(60, 日记账流水号)]),
        MatcherConfig(),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def test_一对多混入另一个明确交易流水号时不得整组自动确认():
    匹配器 = _核对("R2")

    关系 = [
        候选
        for 候选 in 匹配器.candidates
        if 候选.bank_idxs == (0,) and 候选.journal_idxs == (0, 1)
    ]
    assert 关系
    assert all("交易流水号不同" in 候选.evidence.get("business_conflicts", ()) for 候选 in 关系)
    assert not any(
        候选.processing_status
        in {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}
        for 候选 in 匹配器.selected_candidates
    )


def test_一对多各行均为同一交易流水号时可形成明确整组关系():
    匹配器 = _核对("R1")

    assert len(匹配器.selected_candidates) == 1
    候选 = 匹配器.selected_candidates[0]
    assert 候选.bank_idxs == (0,)
    assert 候选.journal_idxs == (0, 1)
    assert 候选.evidence.get("shared_transaction_id") is True
    assert 候选.evidence.get("business_conflicts") == ()
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
