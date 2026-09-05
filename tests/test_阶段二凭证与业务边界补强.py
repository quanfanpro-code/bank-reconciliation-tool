"""第二阶段终审反例：凭证原子性、关键主体同义词和银企批次边界。"""

from decimal import Decimal

import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine


自动通过状态 = {
    ProcessingStatus.AUTO_CONFIRMED,
    ProcessingStatus.GROUP_RECONCILED,
}


def 记录(
    金额,
    摘要,
    *,
    日期="2026-08-10",
    凭证字="",
    凭证号="",
    **辅助字段,
):
    文字字段 = {"摘要": 摘要, **辅助字段}
    return {
        "date": pd.Timestamp(日期),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 文字字段,
        "voucher_word": 凭证字,
        "voucher_no": 凭证号,
    }


def 表格(记录集):
    return pd.DataFrame(
        [
            dict(行, original_idx=序号 + 1, original_file_row=序号 + 2)
            for 序号, 行 in enumerate(记录集)
        ]
    )


def 核对(银行记录, 日记账记录):
    匹配器 = Matcher(
        表格(银行记录),
        表格(日记账记录),
        MatcherConfig(clearly_trivial_threshold=Decimal("5000")),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def test_同一完整凭证内订单号不同也必须保持凭证原子性():
    银行 = [记录(-100, "项目采购付款", 订单号="O-001")]
    日记账 = [
        记录(
            -100,
            "项目采购付款",
            凭证字="记",
            凭证号="001",
            订单号="O-001",
        ),
        记录(
            -200,
            "项目采购付款",
            凭证字="记",
            凭证号="001",
            订单号="O-002",
        ),
    ]

    匹配器 = 核对(银行, 日记账)

    assert frozenset({0, 1}) in 匹配器._atomic_groups["journal"], (
        "年月、凭证字、凭证号和方向相同的日记账行必须先形成完整凭证原子组，"
        "不能被订单号拆开"
    )
    assert not any(
        候选.bank_idxs == (0,)
        and 候选.journal_idxs in {(0,), (1,)}
        and 候选.processing_status in 自动通过状态
        for 候选 in 匹配器.selected_candidates
    ), "银行同额只能命中凭证中的一行时，部分凭证不得自动确认"
    assert any(
        候选.bank_idxs == (0,) and 候选.journal_idxs == (0, 1)
        for 候选 in 匹配器.selected_candidates
    ), "应保留银行一笔对完整凭证两行的关系及其金额差异"


@pytest.mark.parametrize(
    "主体字段",
    ["供应商名称", "收款单位", "付款单位", "对方单位", "对手方", "交易对手"],
)
def test_关键主体字段同义词值不同必须留下冲突证据(主体字段):
    银行 = [记录(-100, "设备采购付款", **{主体字段: "甲公司"})]
    日记账 = [记录(-100, "设备采购付款", **{主体字段: "乙公司"})]

    匹配器 = 核对(银行, 日记账)

    assert len(匹配器.selected_candidates) == 1
    候选 = 匹配器.selected_candidates[0]
    冲突证据 = set(候选.evidence.get("business_conflicts", ()))
    if 候选.text_evidence:
        冲突证据.update(候选.text_evidence.conflicting_fields)
    assert "对方户名" in 冲突证据, f"{主体字段}应统一归入关键主体冲突类别"
    assert 候选.processing_status is ProcessingStatus.FLAGGED
    assert "关键文字字段冲突" in 候选.processing_reason


@pytest.mark.parametrize("批次字段", ["银企批量号", "银企批次号"])
def test_银企批次同义字段必须按两批分别完整核对(批次字段):
    银行 = [
        记录(-40, "供应商批付", **{批次字段: "B-001"}),
        记录(-60, "供应商批付", **{批次字段: "B-001"}),
        记录(-70, "供应商批付", **{批次字段: "B-002"}),
        记录(-130, "供应商批付", **{批次字段: "B-002"}),
    ]
    日记账 = [
        记录(-100, "供应商批付", **{批次字段: "B-001"}),
        记录(-200, "供应商批付", **{批次字段: "B-002"}),
    ]

    匹配器 = 核对(银行, 日记账)

    实际组成 = {
        (候选.bank_idxs, 候选.journal_idxs)
        for 候选 in 匹配器.selected_candidates
    }
    assert 实际组成 == {
        ((0, 1), (0,)),
        ((2, 3), (1,)),
    }, f"{批次字段}必须把B-001和B-002分成两个完整批次"
    assert all(
        候选.processing_status is ProcessingStatus.GROUP_RECONCILED
        and 候选.evidence.get("shared_business_id") is True
        for 候选 in 匹配器.selected_candidates
    )
    assert not any(
        候选.bank_idxs == (0, 1, 2, 3) and 候选.journal_idxs == (0, 1)
        for 候选 in 匹配器.selected_candidates
    ), "同日的两个银企批次不得混成一个4x2关系"
