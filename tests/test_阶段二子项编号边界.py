from decimal import Decimal

import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine
from gui import auto_select_auxiliary_columns


def _记录(金额, **辅助字段):
    return {
        "date": pd.Timestamp("2026-08-10"),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": "设备采购付款",
        "aux_text_fields": {"摘要": "设备采购付款", **辅助字段},
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


def _核对(银行记录, 日记账记录):
    匹配器 = Matcher(
        _表格(银行记录),
        _表格(日记账记录),
        MatcherConfig(),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


@pytest.mark.parametrize("编号字段", ["订单号", "合同号", "结算号"])
def test_一侧额外混入子项编号时不得把整组自动确认(编号字段):
    匹配器 = _核对(
        [_记录(-100, **{编号字段: "A-001"})],
        [
            _记录(-40, **{编号字段: "A-001"}),
            _记录(-60, **{编号字段: "A-002"}),
        ],
    )

    相关候选 = [
        候选
        for 候选 in 匹配器.candidates
        if 候选.bank_idxs == (0,) and 候选.journal_idxs == (0, 1)
    ]
    assert 相关候选
    assert all(候选.evidence.get("shared_business_id") is False for 候选 in 相关候选)
    assert all(
        "业务编号或批次不同" in 候选.evidence.get("business_conflicts", ())
        for 候选 in 相关候选
    )
    assert not any(
        候选.processing_status
        in {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}
        for 候选 in 匹配器.selected_candidates
    )


def test_共同父批次允许双方包含不同订单子项并按完整批次核对():
    匹配器 = _核对(
        [_记录(-100, 批次号="B-001", 订单号="O-001")],
        [
            _记录(-40, 批次号="B-001", 订单号="O-001"),
            _记录(-60, 批次号="B-001", 订单号="O-002"),
        ],
    )

    assert len(匹配器.selected_candidates) == 1
    候选 = 匹配器.selected_candidates[0]
    assert 候选.bank_idxs == (0,)
    assert 候选.journal_idxs == (0, 1)
    assert 候选.evidence.get("shared_business_id") is True
    assert 候选.evidence.get("business_conflicts") == ()
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED


def test_界面默认勾选关键主体和银企批次字段():
    字段 = [
        "供应商名称",
        "收款单位",
        "付款单位",
        "对方单位",
        "对手方",
        "交易对手",
        "银企批量号",
        "银企批次号",
    ]

    assert auto_select_auxiliary_columns(字段) == 字段
