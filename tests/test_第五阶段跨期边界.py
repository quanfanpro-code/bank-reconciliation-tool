"""实际候选生成与报告出口不能遗漏跨期硬边界。"""

import pytest

from data_structures import ProcessingStatus
from reporter import Reporter
from tests.test_单账户业务匹配 import 记录, 核对


@pytest.mark.parametrize("日期", [("2026-01-31", "2026-02-01"), ("2025-12-31", "2026-01-01")])
def test_有共同业务编号的跨月多对多仍为疑点(日期):
    银行日, 账日 = 日期
    m = 核对([记录(-100, "项目结算", "甲公司", "P01", 银行日), 记录(-300, "项目结算", "甲公司", "P01", 银行日)],
             [记录(-150, "项目结算", "甲公司", "P01", 账日), 记录(-250, "项目结算", "甲公司", "P01", 账日)])
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert c.is_cross_month_many_to_many
    assert c.processing_status is ProcessingStatus.FLAGGED
    assert "跨月多对多" in c.processing_reason


@pytest.mark.parametrize("日期", [("2025-12-31", "2026-01-01"), ("2026-01-31", "2026-02-01")])
def test_普通跨期小额自动确认后仍在截止性表完整列示(日期):
    m = 核对([记录(3000, "服务费", 日期=日期[0])], [记录(3000, "服务费", 日期=日期[1])])
    assert m.selected_candidates[0].processing_status is ProcessingStatus.AUTO_CONFIRMED
    表 = Reporter(m).build_report_tables(m.config)["截止性差异"]
    assert len(表) == 1
    assert 表.iloc[0]["匹配ID"] == m.selected_candidates[0].final_match_id
    assert 日期[0] in 表.to_string(index=False)
    assert 日期[1] in 表.to_string(index=False)


def test_同月普通配对不混入截止性表():
    m = 核对([记录(3000, "服务费", 日期="2026-01-01")], [记录(3000, "服务费", 日期="2026-01-02")])
    assert Reporter(m).build_report_tables(m.config)["截止性差异"].empty
