"""第二阶段终审：聚合业务的主体层级差异不能误判为交易对手冲突。"""

from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine


def 记录(
    金额,
    摘要,
    对方,
    *,
    日期="2026-08-10",
    凭证字="",
    凭证号="",
    批次号="",
):
    辅助字段 = {"摘要": 摘要, "对方户名": 对方}
    if 批次号:
        辅助字段["批次号"] = 批次号
    return {
        "date": pd.Timestamp(日期),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 辅助字段,
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
        MatcherConfig(),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def 找到完整关系(匹配器, 银行索引, 日记账索引):
    符合 = [
        候选
        for 候选 in 匹配器.selected_candidates
        if 候选.bank_idxs == 银行索引 and 候选.journal_idxs == 日记账索引
    ]
    assert len(符合) == 1, "应选中覆盖双方完整范围的唯一关系"
    return 符合[0]


def 断言只是聚合主体层级差异(候选):
    assert "对方户名" not in set(候选.evidence.get("business_conflicts", ()))
    if 候选.text_evidence:
        assert "对方户名" not in set(候选.text_evidence.conflicting_fields)
    提示 = str(候选.evidence.get("aggregate_party_hierarchy_hint", ""))
    assert "主体" in 提示 and "汇总" in 提示 and "明细" in 提示, (
        "平台或汇总主体与员工、供应商明细主体不同，应作为聚合主体层级提示留痕"
    )
    assert "层级不同" in str(候选.evidence.get("batch_review_hint", "")), (
        "报告使用的批次核查线索必须说明汇总与明细层级不同"
    )


def test_工资批扣平台对同凭证35笔员工明细应按完整组闭合():
    银行 = [记录(-3500, "2026年8月工资代发批扣", "工资代发平台")]
    日记账 = [
        记录(
            -100,
            "2026年8月工资",
            f"员工{序号:02d}",
            凭证字="记",
            凭证号="0088",
        )
        for 序号 in range(1, 36)
    ]

    匹配器 = 核对(银行, 日记账)

    候选 = 找到完整关系(匹配器, (0,), tuple(range(35)))
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
    assert 候选.evidence.get("salary_group") is True
    assert 候选.evidence.get("resolves_full_group") is True
    断言只是聚合主体层级差异(候选)


def test_共同批次号的银行供应商批付汇总对多供应商明细应完整闭合():
    批次 = "PAY-20260810-001"
    银行 = [记录(-1000, "供应商批付汇总", "银企批付平台", 批次号=批次)]
    日记账 = [
        记录(-200, "供应商批付", "甲供应商", 批次号=批次),
        记录(-300, "供应商批付", "乙供应商", 批次号=批次),
        记录(-500, "供应商批付", "丙供应商", 批次号=批次),
    ]

    匹配器 = 核对(银行, 日记账)

    候选 = 找到完整关系(匹配器, (0,), (0, 1, 2))
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
    assert 候选.evidence.get("shared_business_id") is True
    assert 候选.evidence.get("shared_business_category") == "批次"
    assert 候选.evidence.get("resolves_full_group") is True
    断言只是聚合主体层级差异(候选)
    线索 = str(候选.evidence.get("batch_review_hint", ""))
    assert "供应商付款清单" in 线索
    assert "工资表" not in 线索


def test_完整工资凭证同时有金额差异时两类核查线索都必须保留():
    银行 = [记录(-900, "2026年8月工资代发批扣", "工资代发平台")]
    日记账 = [
        记录(
            -100,
            "2026年8月工资",
            f"员工{序号:02d}",
            凭证字="记",
            凭证号="0099",
        )
        for 序号 in range(1, 11)
    ]

    匹配器 = 核对(银行, 日记账)
    候选 = 找到完整关系(匹配器, (0,), tuple(range(10)))

    线索 = str(候选.evidence.get("batch_review_hint", ""))
    assert "完整凭证与银行记录存在金额差异" in 线索
    assert "汇总主体" in 线索 and "明细主体" in 线索 and "层级不同" in 线索


def test_普通一对一同额货款的明确主体不同仍是硬冲突():
    匹配器 = 核对(
        [记录(-1000, "采购货款", "甲供应商")],
        [记录(-1000, "采购货款", "乙供应商")],
    )

    候选 = 找到完整关系(匹配器, (0,), (0,))
    冲突 = set(候选.evidence.get("business_conflicts", ()))
    if 候选.text_evidence:
        冲突.update(候选.text_evidence.conflicting_fields)
    assert "对方户名" in 冲突
    assert 候选.processing_status is ProcessingStatus.FLAGGED
    assert "关键文字字段冲突" in 候选.processing_reason
    assert not 候选.evidence.get("aggregate_party_hierarchy_hint")
