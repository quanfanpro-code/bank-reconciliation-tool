"""独立业务样例：费用方向、明确编号冲突和原行占用。"""

import pytest

from data_structures import MatcherConfig
from matcher import Matcher
from tests.test_第三阶段特殊业务 import _df
from 业务事件 import detect_same_side_events


@pytest.mark.parametrize("费用来源", ["journal", "bank"])
def test_付款和手续费同为支出仍能核对完整净额(费用来源):
    明细 = _df([("2026-01-05", -1000, {"摘要": "设备款", "业务编号": "P01"}),
                 ("2026-01-05", -10, {"摘要": "银行手续费", "业务编号": "P01"})])
    汇总 = _df([("2026-01-05", -1010, {"摘要": "设备款及手续费", "业务编号": "P01"})])
    m = Matcher(汇总 if 费用来源 == "journal" else 明细,
                明细 if 费用来源 == "journal" else 汇总, MatcherConfig())
    m.run()
    结果 = [项 for 项 in m.candidates if 项.match_type == "fee_net"]
    assert len(结果) == 1
    assert 结果[0].metrics.total_diff_li == 0
    assert 结果[0].evidence["fee_amount_li"] == 100000


@pytest.mark.parametrize("冲突字段,值", [("业务编号", "P02"), ("对方", "乙公司")])
def test_同侧退款不能无视明确业务冲突(冲突字段, 值):
    原字段 = {"摘要": "支付货款", "业务编号": "P01", "对方": "甲公司"}
    退款字段 = {**原字段, "摘要": "退款", 冲突字段: 值}
    frame = _df([("2026-01-02", -100, 原字段), ("2026-01-03", 100, 退款字段)])
    事件, 线索 = detect_same_side_events(frame, "bank")
    assert 事件 == []
    assert 线索 and len(frame) == 2


def test_两次退回重付不能重复占用上一条业务链原行():
    rows = [(f"2026-01-0{day}", amount, {"摘要": summary, "业务编号": "P01", "对方": "甲公司"})
            for day, amount, summary in [(1, -100, "付款"), (2, 100, "退汇"),
                                         (3, -100, "重付"), (4, 100, "退汇"), (5, -100, "重付")]]
    事件, _ = detect_same_side_events(_df(rows), "bank")
    原行 = [row for event in 事件 for row in event.row_idxs]
    assert len(原行) == len(set(原行))


def test_零金额记录不能制造正负同额退款线索():
    frame = _df([("2026-01-02", 0, {"摘要": "查询"}),
                 ("2026-01-03", 0, {"摘要": "查询"})])
    事件, 线索 = detect_same_side_events(frame, "bank")
    assert not 事件
    assert not any(项.clue_type == "正负同额线索" for 项 in 线索)
