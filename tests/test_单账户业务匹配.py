"""通过真实匹配入口验证业务组成，防止只凭金额凑平。"""
from decimal import Decimal
from time import perf_counter

import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine


def 记录(金额, 摘要="货款", 对方="", 编号="", 日期="2026-08-10", 凭证=""):
    字段 = {"摘要": 摘要}
    if 对方:
        字段["对方户名"] = 对方
    if 编号:
        字段["业务编号"] = 编号
    return {"date": pd.Timestamp(日期), "amount": Decimal(str(金额)),
            "amount_decimal": PrecisionEngine.to_integer_li(金额),
            "summary": 摘要, "aux_text_fields": 字段, "voucher_no": 凭证}


def 核对(银行, 账, **配置):
    def 表(行):
        return pd.DataFrame([dict(r, original_idx=i + 1, original_file_row=i + 2)
                             for i, r in enumerate(行)])
    匹配器 = Matcher(表(银行), 表(账), MatcherConfig(**配置), logger=lambda _: None)
    匹配器.run()
    return 匹配器


@pytest.mark.parametrize("反向", [False, True])
def test_不相邻同业务完整拆分优先于局部差额(反向):
    单 = [记录(-1000, "支付货款", "甲公司", "P01")]
    多 = [记录(-400, "支付货款", "甲公司", "P01"),
          记录(-200, "材料款", "乙公司", "P99"),
          记录(-600, "支付货款", "甲公司", "P01")]
    m = 核对(多, 单) if 反向 else 核对(单, 多)
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert (c.bank_idxs, c.journal_idxs) == (((0, 2), (0,)) if 反向 else ((0,), (0, 2)))
    assert c.processing_status in {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}
    assert c.evidence["business_basis"]
    assert not c.is_ambiguous


def test_错误同额组合不能污染正确业务歧义():
    m = 核对([记录(-200, "货款", "甲公司", "P02"), 记录(-800, "货款", "甲公司", "P02"),
              记录(-400, "材料", "乙公司", "P88"), 记录(-600, "材料", "乙公司", "P88")],
             [记录(-1000, "货款", "甲公司", "P02")])
    c = m.selected_candidates[0]
    assert c.bank_idxs == (0, 1)
    assert not c.is_ambiguous


@pytest.mark.parametrize("编号", ["G08", ""])
def test_不等额工资完整归组且不挪用到材料款(编号):
    工资 = [记录(-(3000 + i * 100), "8月工资", f"员工{i + 1}", 编号) for i in range(35)]
    工资.insert(17, 记录(-9000, "设备款", "设备供应商", "P90"))
    m = 核对(工资, [记录(-164500, "8月工资", "工资", 编号), 记录(-17000, "材料款", "材料供应商", "P91")])
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert c.bank_idxs == tuple(i for i in range(36) if i != 17)
    assert c.journal_idxs == (0,)
    assert c.processing_status in {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}


def test_同业务两对三完整对应():
    m = 核对([记录(-150, "结算", "甲公司", "P03"), 记录(-450, "结算", "甲公司", "P03"), 记录(-900, "设备", "乙公司", "P90")],
             [记录(-100, "结算", "甲公司", "P03"), 记录(-200, "结算", "甲公司", "P03"), 记录(-300, "结算", "甲公司", "P03"), 记录(-1700, "材料", "丙公司", "P91")])
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert (c.bank_idxs, c.journal_idxs) == ((0, 1), (0, 1, 2))
    assert c.processing_status is ProcessingStatus.GROUP_RECONCILED


def test_业务编号一致的差额优先于无依据同额():
    m = 核对([记录(-1000, "项目款", "", "P01")],
             [记录(-990, "项目款", "", "P01"), 记录(-1000, "", "")])
    c = m.selected_candidates[0]
    assert c.journal_idxs == (0,)
    assert c.metrics.total_diff_li == PrecisionEngine.to_integer_li(10)
    assert c.processing_status is ProcessingStatus.AUTO_CLASSIFIED


def test_不同明示批次不能借宽泛摘要重新合并():
    m = 核对([记录(-1000, "8月工资 批次A")],
             [记录(-400, "8月工资 批次A"), 记录(-600, "8月工资 批次B")])
    assert not any(c.journal_idxs == (0, 1) and c.processing_status in
                   {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}
                   for c in m.selected_candidates)


def test_不同期间同凭证号不能作为一个业务组():
    m = 核对([记录(-1000, "项目款", 日期="2026-02-01")],
             [记录(-400, "项目款", 日期="2026-01-31", 凭证="记001"),
              记录(-600, "项目款", 日期="2026-02-01", 凭证="记001")])
    assert not any(c.journal_idxs == (0, 1) and c.processing_status in
                   {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}
                   for c in m.selected_candidates)


def test_同额两对一仍有竞争组成():
    m = 核对([记录(100, "收款", "甲公司"), 记录(100, "收款", "甲公司")], [记录(100, "收款", "甲公司")])
    c = m.selected_candidates[0]
    assert c.processing_status is ProcessingStatus.FLAGGED
    assert c.evidence["alternative_candidate_ids"]


def test_同额两对两保留整组一致():
    行 = [记录(100, "收款", "甲公司"), 记录(100, "收款", "甲公司")]
    m = 核对(行, 行)
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert (c.bank_idxs, c.journal_idxs) == ((0, 1), (0, 1))
    assert c.processing_status is ProcessingStatus.GROUP_RECONCILED


def test_工资双方实际主体不同仍是冲突():
    m = 核对([记录(-100, "8月工资", "张三")], [记录(-100, "8月工资", "李四")])
    assert all(c.processing_status is ProcessingStatus.FLAGGED for c in m.selected_candidates)


def test_整组标记不能绕过明确主体冲突():
    m = 核对([记录(-40, "货款", "甲公司"), 记录(-60, "货款", "甲公司")],
             [记录(-100, "货款", "乙公司")])
    assert not any(c.processing_status in {ProcessingStatus.AUTO_CONFIRMED, ProcessingStatus.GROUP_RECONCILED}
                   for c in m.selected_candidates)


def test_业务编号与批次号是不同字段且可同时提供():
    银行 = [记录(-400, "8月工资", "员工甲", "PAY08"), 记录(-600, "8月工资", "员工乙", "PAY08")]
    账 = [记录(-1000, "2026年8月工资", "工资", "PAY08")]
    for 行 in 银行 + 账:
        行["aux_text_fields"]["批次号"] = "G08"
    m = 核对(银行, 账)
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert (c.bank_idxs, c.journal_idxs) == ((0, 1), (0,))
    assert c.processing_status is ProcessingStatus.GROUP_RECONCILED


def test_完整工资组不再进入通用子集搜索且记录覆盖范围():
    m = 核对([记录(-100, "8月工资", f"员工{i}", "G08") for i in range(40)],
             [记录(-4000, "8月工资", "工资", "G08")])
    覆盖 = m.run_parameters["combination_search"]
    assert 覆盖["business_group_bank_rows"] == 40
    assert 覆盖["business_group_journal_rows"] == 1
    assert 覆盖["generic_source_rows"] == 0
    assert m.selected_candidates[0].bank_idxs == tuple(range(40))


def test_同日同额两批工资对无批次汇总保留整批竞争():
    银行 = [记录(-400, "8月工资", "张三", "A"), 记录(-600, "8月工资", "李四", "A"),
            记录(-300, "8月工资", "王五", "B"), 记录(-700, "8月工资", "赵六", "B")]
    m = 核对(银行, [记录(-1000, "8月工资", "工资")])
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert c.processing_status is ProcessingStatus.FLAGGED
    assert c.bank_idxs in {(0, 1), (2, 3)}
    alternatives = {a.candidate_id: a for a in m.candidates}
    assert any(alternatives[i].bank_idxs == ((2, 3) if c.bank_idxs == (0, 1) else (0, 1))
               for i in c.evidence["alternative_candidate_ids"])


def test_两个业务各差十元不能以合计相等抵销():
    m = 核对([记录(-1000, "项目款", 编号="A"), 记录(-1000, "项目款", 编号="B")],
             [记录(-990, "项目款", 编号="A"), 记录(-1010, "项目款", 编号="B")])
    assert [(c.bank_idxs, c.journal_idxs) for c in m.selected_candidates] == [((0,), (0,)), ((1,), (1,))]
    assert all(c.metrics.total_diff_li == PrecisionEngine.to_integer_li(10)
               and c.processing_status is ProcessingStatus.AUTO_CLASSIFIED for c in m.selected_candidates)


def test_数字业务编号前导零仍参与区分():
    m = 核对([记录(-1000, "项目款", 编号="001")],
             [记录(-1000, "项目款", 编号="1"), 记录(-1000, "项目款", 编号="001")])
    c = m.selected_candidates[0]
    assert c.journal_idxs == (1,)
    assert c.processing_status is ProcessingStatus.AUTO_CONFIRMED


def test_用户取消后不提交已生成候选():
    frame = pd.DataFrame([dict(记录(-100, "8月工资", 编号="G08"), original_idx=1)])
    m = Matcher(frame, frame, MatcherConfig())
    def 日志(消息):
        if 消息.startswith("开始: 智能组合"):
            m.set_stopping(True)
    m.logger = 日志
    assert m.run() == []
    assert not m.bank["matched"].any()
    assert not m.journal["matched"].any()


def test_大量同额子集保留真实多解且不展开全部排列():
    开始 = perf_counter()
    m = 核对([记录(-12, "")], [记录(-1, "") for _ in range(24)],
             clearly_trivial_threshold=Decimal("0"))
    assert m.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED
    assert len(m.selected_candidates[0].journal_idxs) == 12
    assert perf_counter() - 开始 < 3


@pytest.mark.parametrize("摘要", ["", "货款"])
def test_不同明确对方的逐笔关系不能被日总额组吞并(摘要):
    行 = [记录(-100, 摘要, "甲公司"), 记录(-200, 摘要, "乙公司")]
    m = 核对(行, 行)
    assert {(c.bank_idxs, c.journal_idxs) for c in m.selected_candidates} == {((0,), (0,)), ((1,), (1,))}
    assert all(c.processing_status is ProcessingStatus.AUTO_CONFIRMED for c in m.selected_candidates)


@pytest.mark.parametrize("工资", [False, True])
def test_强业务依据完整组保留差额而不是局部差额(工资):
    if 工资:
        银行 = [记录(-(3000 + i * 100), "8月工资", f"员工{i + 1}", "G08") for i in range(35)]
        账 = [记录(-164490, "8月工资", "工资", "G08")]
        期望银行, 期望账 = tuple(range(35)), (0,)
    else:
        银行 = [记录(-150, "项目款", "甲公司", "P03"), 记录(-450, "项目款", "甲公司", "P03")]
        账 = [记录(-100, "项目款", "甲公司", "P03"), 记录(-200, "项目款", "甲公司", "P03"), 记录(-290, "项目款", "甲公司", "P03")]
        期望银行, 期望账 = (0, 1), (0, 1, 2)
    m = 核对(银行, 账)
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert (c.bank_idxs, c.journal_idxs) == (期望银行, 期望账)
    assert c.metrics.total_diff_li == PrecisionEngine.to_integer_li(10)
    assert c.processing_status is ProcessingStatus.AUTO_CLASSIFIED
    assert not c.is_ambiguous


def test_同对方不同具体用途的差额不能被合计抵销():
    m = 核对([记录(-1000, "采购设备", "甲公司"), 记录(-2000, "采购材料", "甲公司")],
             [记录(-990, "采购设备", "甲公司"), 记录(-2010, "采购材料", "甲公司")])
    assert {(c.bank_idxs, c.journal_idxs) for c in m.selected_candidates} == {((0,), (0,)), ((1,), (1,))}
    assert all(c.metrics.total_diff_li == PrecisionEngine.to_integer_li(10)
               and c.processing_status is ProcessingStatus.AUTO_CLASSIFIED for c in m.selected_candidates)


def test_跨年工资未明示年份时不能凭交易年制造冲突():
    m = 核对([记录(-400, "12月工资", "张三", "G2512", "2026-01-01"),
              记录(-600, "12月工资", "李四", "G2512", "2026-01-01")],
             [记录(-1000, "12月工资", "工资", "G2512", "2025-12-31")])
    c = m.selected_candidates[0]
    assert (c.bank_idxs, c.journal_idxs) == ((0, 1), (0,))
    assert c.processing_status is ProcessingStatus.GROUP_RECONCILED


@pytest.mark.parametrize("摘要", ["转账", "银行转账", "转账支出", "网上转账"])
def test_宽泛转账摘要不能排除同额竞争(摘要):
    m = 核对([记录(-1000, 摘要)], [记录(-990, 摘要), 记录(-1000, "支付货款")])
    c = m.selected_candidates[0]
    assert c.journal_idxs == (1,)
    assert c.processing_status is ProcessingStatus.FLAGGED
    assert c.evidence["alternative_candidate_ids"]


def test_业务编号不同分隔符不能被当作同号():
    m = 核对([记录(-1000, "项目款", 编号="AB-12")],
             [记录(-1000, "项目款", 编号="AB12")])
    assert m.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED


def test_工资明确写明不同年份仍保留冲突():
    m = 核对([记录(-1000, "2025年12月工资", 编号="G12", 日期="2026-01-01")],
             [记录(-1000, "2026年12月工资", 编号="G12", 日期="2026-01-01")])
    assert m.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED
