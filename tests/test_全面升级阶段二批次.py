"""第二阶段：完整批次、凭证组与等价解释的验收行为。"""
from decimal import Decimal

import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter
from 业务分组 import complete_groups, row_business


def 记录(金额, 摘要, 对方="", 日期="2026-08-10", 凭证=""):
    字段 = {"摘要": 摘要}
    if 对方:
        字段["对方户名"] = 对方
    return {
        "date": pd.Timestamp(日期),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 字段,
        "voucher_no": 凭证,
    }


def 表格(记录集):
    return pd.DataFrame(
        [
            dict(行, original_idx=序号 + 1, original_file_row=序号 + 2)
            for 序号, 行 in enumerate(记录集)
        ]
    )


def 核对(银行, 序时账, **配置):
    参数 = {"clearly_trivial_threshold": Decimal("5000"), **配置}
    匹配器 = Matcher(
        表格(银行),
        表格(序时账),
        MatcherConfig(**参数),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def test_同一完整凭证35行摘要不同仍独立组成一笔银行付款():
    银行 = [
        记录(-3500, "甲公司项目结算", "甲公司"),
        记录(-200, "购买办公设备", "乙公司"),
    ]
    序时账 = [
        记录(-100, f"项目明细{i:02d}", "甲公司", 凭证="记001")
        for i in range(1, 36)
    ] + [记录(-200, "购买办公设备", "乙公司", 凭证="记002")]

    匹配器 = 核对(银行, 序时账)

    凭证关系 = next(
        (
            候选
            for 候选 in 匹配器.selected_candidates
            if 候选.bank_idxs == (0,) and 候选.journal_idxs == tuple(range(35))
        ),
        None,
    )
    assert 凭证关系 is not None
    assert 凭证关系.processing_status is ProcessingStatus.GROUP_RECONCILED
    assert 凭证关系.evidence.get("resolves_full_group") is True
    assert "凭证" in 凭证关系.evidence.get("business_basis", "")
    assert any(
        候选.bank_idxs == (1,) and 候选.journal_idxs == (35,)
        for 候选 in 匹配器.selected_candidates
    )


def test_无编号35笔工资与账面整笔差10元仍保留一个完整批次差异():
    银行 = [记录(-100, "8月工资", f"员工{i:02d}") for i in range(1, 36)]
    序时账 = [记录(-3490, "8月工资汇总", "工资")]

    匹配器 = 核对(银行, 序时账)

    assert len(匹配器.selected_candidates) == 1
    批次关系 = 匹配器.selected_candidates[0]
    assert 批次关系.bank_idxs == tuple(range(35))
    assert 批次关系.journal_idxs == (0,)
    assert 批次关系.metrics.total_diff_li == PrecisionEngine.to_integer_li(10)
    assert 批次关系.processing_status is ProcessingStatus.AUTO_CLASSIFIED
    assert 批次关系.evidence.get("resolves_full_group") is True


def test_同额逐笔与拆分两套解释换序后都披露为多解():
    银行 = [记录(-1000, "甲公司项目结算", "甲公司")]
    原序 = [
        记录(-1000, "甲公司项目结算", "甲公司"),
        记录(-400, "甲公司项目结算", "甲公司"),
        记录(-600, "甲公司项目结算", "甲公司"),
    ]
    换序 = [原序[2], 原序[0], 原序[1]]

    def 等额解释(匹配器):
        return {
            tuple(sorted(匹配器.journal.at[i, "amount"] for i in 候选.journal_idxs))
            for 候选 in 匹配器.candidates
            if 候选.bank_idxs == (0,)
            and 候选.metrics.total_diff_li == 0
            and not 候选.evidence.get("business_conflicts")
        }

    两次结果 = [核对(银行, 原序), 核对(银行, 换序)]
    手工预期 = {
        (Decimal("-1000"),),
        (Decimal("-600"), Decimal("-400")),
    }

    assert 等额解释(两次结果[0]) == 手工预期
    assert 等额解释(两次结果[1]) == 手工预期
    for 匹配器 in 两次结果:
        assert len(匹配器.selected_candidates) == 1
        最终关系 = 匹配器.selected_candidates[0]
        assert 最终关系.processing_status is ProcessingStatus.FLAGGED
        assert 最终关系.is_ambiguous is True
        assert 最终关系.evidence.get("alternative_candidate_ids")


def test_同一凭证号跨不同完整年月不能标成完整凭证组():
    银行 = [记录(-1000, "甲公司年度结算", "甲公司", 日期="2026-08-31")]
    序时账 = [
        记录(-400, "上年结算明细", "甲公司", 日期="2025-08-31", 凭证="记001"),
        记录(-600, "本年结算明细", "甲公司", 日期="2026-08-31", 凭证="记001"),
    ]
    业务行 = {
        i: row_business(行)
        for i, 行 in 表格(序时账).iterrows()
    }

    凭证组 = [组 for 组 in complete_groups(业务行, 400) if 组[0][0] == "凭证"]
    assert [组[0][1] for 组 in 凭证组] == ["2025-08", "2026-08"]
    assert [组[1] for 组 in 凭证组] == [(0,), (1,)]

    匹配器 = 核对(银行, 序时账, dfs_date_window=400, tolerance_days=400)
    assert not any(
        候选.match_type == "business_group"
        and 候选.evidence.get("business_group_kind") == "凭证"
        and 候选.journal_idxs == (0, 1)
        for 候选 in 匹配器.candidates
    )


def test_同期间完整凭证超窗口仍保持原子组且整组拒绝():
    银行 = [记录(-1500, "甲公司项目结算", "甲公司", 日期="2026-08-01")]
    序时账 = [
        记录(-400, "材料明细", "甲公司", 日期="2026-08-01", 凭证="记003"),
        记录(-600, "人工明细", "甲公司", 日期="2026-08-02", 凭证="记003"),
        记录(-500, "运输明细", "甲公司", 日期="2026-08-20", 凭证="记003"),
    ]
    业务行 = {
        i: row_business(行)
        for i, 行 in 表格(序时账).iterrows()
    }

    凭证组 = [组 for 组 in complete_groups(业务行, 3) if 组[0][0] == "凭证"]
    assert len(凭证组) == 1
    assert 凭证组[0][1] == (0, 1, 2)

    匹配器 = 核对(银行, 序时账, dfs_date_window=3, tolerance_days=3)
    assert not any(
        候选.match_type == "business_group"
        and 候选.evidence.get("business_group_kind") == "凭证"
        for 候选 in 匹配器.candidates
    )


def test_银行与完整凭证对方明确冲突不得自动形成凭证关系():
    银行 = [记录(-1000, "项目结算", "甲公司")]
    序时账 = [
        记录(-400, "项目材料", "乙公司", 凭证="记005"),
        记录(-600, "项目人工", "乙公司", 凭证="记005"),
    ]

    匹配器 = 核对(银行, 序时账)

    assert not any(
        候选.match_type == "business_group"
        and 候选.evidence.get("business_group_kind") == "凭证"
        for 候选 in 匹配器.candidates
    )
    完整关系 = [
        候选
        for 候选 in 匹配器.selected_candidates
        if 候选.bank_idxs == (0,) and 候选.journal_idxs == (0, 1)
    ]
    assert all(
        候选.processing_status is ProcessingStatus.FLAGGED
        and "对方户名" in 候选.evidence.get("business_conflicts", ())
        for 候选 in 完整关系
    )


@pytest.mark.parametrize(
    ("业务摘要", "汇总摘要", "方向", "各方"),
    [
        ("员工报销", "员工报销汇总", -1, ("张三", "李四", "王五", "赵六")),
        (
            "供应商批量付款",
            "供应商批量付款汇总",
            -1,
            ("供应商甲", "供应商乙", "供应商丙", "供应商丁"),
        ),
        (
            "客户批量收款",
            "客户批量收款汇总",
            1,
            ("客户甲", "客户乙", "客户丙", "客户丁"),
        ),
    ],
)
def test_无编号明确批量业务按用途完整归组(业务摘要, 汇总摘要, 方向, 各方):
    金额 = [120, 230, 340, 410]
    银行 = [
        记录(方向 * 金额值, 业务摘要, 对方)
        for 金额值, 对方 in zip(金额, 各方)
    ]
    序时账 = [记录(方向 * 1100, 汇总摘要)]

    匹配器 = 核对(银行, 序时账)

    assert len(匹配器.selected_candidates) == 1
    批量关系 = 匹配器.selected_candidates[0]
    assert 批量关系.bank_idxs == (0, 1, 2, 3)
    assert 批量关系.journal_idxs == (0,)
    assert 批量关系.processing_status is ProcessingStatus.GROUP_RECONCILED
    assert 批量关系.evidence.get("resolves_full_group") is True
    assert 业务摘要.replace("批量付款", "").replace("批量收款", "")[:2] in 批量关系.evidence.get(
        "business_basis", ""
    )


@pytest.mark.parametrize("银行笔数", [34, 36])
def test_工资批次不完整时证据和报告列明组成差额及核查线索(银行笔数):
    银行 = [记录(-100, "8月工资", f"员工{i:02d}") for i in range(1, 银行笔数 + 1)]
    序时账 = [记录(-3500, "8月工资汇总", "工资")]

    匹配器 = 核对(银行, 序时账)

    assert len(匹配器.selected_candidates) == 1
    批次关系 = 匹配器.selected_candidates[0]
    assert len(批次关系.bank_idxs) == 银行笔数
    assert 批次关系.journal_idxs == (0,)
    assert 批次关系.metrics.total_diff_li == PrecisionEngine.to_integer_li(100)
    assert 批次关系.evidence.get("batch_difference") is True
    assert 批次关系.evidence.get("batch_bank_count") == 银行笔数
    assert 批次关系.evidence.get("batch_journal_count") == 1
    assert 批次关系.evidence.get("batch_difference_li") == PrecisionEngine.to_integer_li(100)
    核查线索 = 批次关系.evidence.get("batch_review_hint", "")
    assert "缺项" in 核查线索 or "重复" in 核查线索

    报表 = Reporter(匹配器).build_report_tables(匹配器.config)
    整组 = 报表["整组勾稽"].loc[
        报表["整组勾稽"]["匹配ID"] == 批次关系.final_match_id
    ]
    事项 = 报表["疑点事项"].loc[
        报表["疑点事项"]["匹配ID"] == 批次关系.final_match_id
    ]
    assert len(整组) == len(事项) == 1
    assert 整组.iloc[0]["银行笔数"] == 事项.iloc[0]["银行笔数"] == 银行笔数
    assert 整组.iloc[0]["日记账笔数"] == 事项.iloc[0]["日记账笔数"] == 1
    assert 整组.iloc[0]["总差额"] == 事项.iloc[0]["差异金额"] == 100
    for 行 in (整组.iloc[0], 事项.iloc[0]):
        报表文字 = "；".join(str(值) for 值 in 行.tolist() if pd.notna(值))
        assert "缺项" in 报表文字 or "重复" in 报表文字
    assert 报表["银行侧待查"].empty
    assert 报表["日记账侧待查"].empty


def test_普通货款不能仅凭同日同总额跨不同对方强行整组():
    银行 = [
        记录(-400, "支付货款", "甲公司"),
        记录(-600, "支付货款", "乙公司"),
    ]
    序时账 = [
        记录(-500, "支付货款", "甲公司"),
        记录(-500, "支付货款", "乙公司"),
    ]

    匹配器 = 核对(银行, 序时账)

    强行整组 = [
        候选
        for 候选 in 匹配器.selected_candidates
        if 候选.bank_idxs == (0, 1) and 候选.journal_idxs == (0, 1)
    ]
    assert all(
        候选.processing_status is ProcessingStatus.FLAGGED
        for 候选 in 强行整组
    )
