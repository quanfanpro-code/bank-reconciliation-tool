"""完整原始总体的复核分流，不把程序已确认的大额重新推给人工。"""

from copy import deepcopy
from decimal import Decimal
from types import SimpleNamespace

import pandas as pd
import pytest

from data_structures import BusinessEvent, MatcherConfig, OverallControlResult
from matcher import Matcher
from precision_engine import PrecisionEngine
from 复核事项 import build_review_items


def 记录(rows):
    return pd.DataFrame([
        {"date": pd.Timestamp(day), "amount": Decimal(str(amount)),
         "amount_decimal": PrecisionEngine.to_integer_li(amount),
         "summary": f"业务{index}", "aux_text_fields": {"对方户名": "甲公司", "业务编号": f"P{index}"},
         "original_idx": index + 1, "original_file_row": index + 2,
         "voucher_no": "", "balance": None}
        for index, (day, amount) in enumerate(rows)
    ], columns=["date", "amount", "amount_decimal", "summary", "aux_text_fields", "original_idx", "original_file_row", "voucher_no", "balance"])


def 未匹配(bank, journal=(), control=None):
    return SimpleNamespace(bank=记录(bank), journal=记录(journal), config=MatcherConfig(),
                           selected_candidates=[], difference_pools=[], business_events=[],
                           overall_control=control, overall_scope_limited=bool(control and control.scope_limited))


def test_二十万明确对应仍自动确认且不改原匹配器():
    bank = 记录([("2026-01-02", 200000)])
    matcher = Matcher(bank, bank.copy(deep=True), MatcherConfig())
    matcher.run()
    before_bank = matcher.bank.copy(deep=True)
    before_candidates = deepcopy(matcher.selected_candidates)
    items = build_review_items(matcher)
    assert len(items) == 1
    assert items[0]["核对方式"] == "自动确认"
    assert items[0]["影响金额"] == Decimal("0")
    assert items[0]["事项编号"] == matcher.selected_candidates[0].final_match_id
    pd.testing.assert_frame_equal(before_bank, matcher.bank)
    assert matcher.selected_candidates == before_candidates


def test_二十万单边全查不再无差别写范围未知():
    item = build_review_items(未匹配([("2026-01-02", -200000)]))[0]
    assert (item["风险等级"], item["核对方式"]) == ("高风险", "人工全查")
    assert item["影响金额"] == item["差异金额"] == Decimal("200000")
    assert item["银行支出"] == Decimal("200000")
    assert item["银行索引"] == (0,) and item["日记账索引"] == ()


def test_同月二十一笔五千单边累计全部升级():
    items = build_review_items(未匹配([("2026-01-02", 5000)] * 21))
    assert len(items) == 21
    assert all(item["核对方式"] == "人工全查" for item in items)
    assert all("月度累计" in item["入选原因"] for item in items)


@pytest.mark.parametrize("rows, journal", [
    ([("2026-01-02", 60000), ("2026-02-02", 60000)], ()),
    ([("2025-01-02", 60000), ("2026-01-02", 60000)], ()),
    ([("2026-01-02", 60000), ("2026-01-02", -60000)], ()),
    ([("2026-01-02", 60000)], [("2026-01-02", 60000)]),
])
def test_不同年月来源收支方向分别累计(rows, journal):
    items = build_review_items(未匹配(rows, journal))
    assert all(item["风险等级"] == "中风险" for item in items)
    assert all(item["核对方式"] == "留存备查" for item in items)


def test_完整总体中风险按最早日期等距抽取():
    matcher = 未匹配([("2026-03-04", 6000), ("2026-03-01", 6000),
                      ("2026-03-03", 6000), ("2026-03-02", 6000), ("2026-04-01", 200000)])
    items = build_review_items(matcher)
    sampled = [item for item in items if item["核对方式"] == "人工抽样"]
    assert len(sampled) == 1
    assert sampled[0]["首次日期"] == pd.Timestamp("2026-03-01")
    assert sum(item["核对方式"] == "留存备查" for item in items) == 3


def test_真实断档仅涉及窗口记录范围未知且不扩张抽样基数():
    control = OverallControlResult(affected_windows=((pd.Timestamp("2026-01-02"), pd.Timestamp("2026-01-03")),))
    items = build_review_items(未匹配([("2026-01-02", 200000), ("2026-02-04", 6000)], control=control))
    assert items[0]["风险等级"] == "范围未知"
    assert items[0]["核对方式"] == "人工全查"
    assert items[1]["风险等级"] == "中风险"
    assert items[1]["核对方式"] == "留存备查"


def test_全部原记录唯一归属且完整同侧链保留():
    matcher = 未匹配([("2025-12-31", -3000), ("2026-01-01", 3000), ("2026-01-02", -3000)], [("2026-01-04", 12)])
    matcher.business_events = [BusinessEvent("E1", "bank", "退回重付", (0, 1, 2), -30000000, "原付款＋退回＋重付", "已有交易编号", True)]
    items = build_review_items(matcher)
    assert len(items) == 2
    chain = next(item for item in items if item["银行索引"])
    assert chain["程序结论"] == "同侧业务链"
    assert chain["银行索引"] == (0, 1, 2)
    assert chain["影响金额"] == Decimal("9000")
    assert chain["跨期"] is True
    assert sorted(index for item in items for index in item["银行索引"]) == [0, 1, 2]
    assert sorted(index for item in items for index in item["日记账索引"]) == [0]
    assert len({item["事项编号"] for item in items}) == len(items)


def test_单边与原差异池同方向合并后先升级再抽样():
    bank = 记录([("2026-01-02", 70000)])
    journal = 记录([("2026-01-02", 10000)])
    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()
    assert len(matcher.difference_pools) == 1
    matcher.bank.loc[1] = 记录([("2026-01-03", 50000)]).iloc[0]
    matcher.bank.loc[1, "original_file_row"] = 3
    before_risk = matcher.selected_candidates[0].risk_level
    items = build_review_items(matcher)
    assert len(items) == 2
    assert all(item["核对方式"] == "人工全查" for item in items)
    assert matcher.selected_candidates[0].risk_level == before_risk


def test_单边事项编号不依赖显示顺序():
    matcher = 未匹配([("2026-01-02", 10), ("2026-01-03", 20)])
    original = {item["银行索引"]: item["事项编号"] for item in build_review_items(matcher)}
    matcher.bank = matcher.bank.iloc[::-1]
    assert {item["银行索引"]: item["事项编号"] for item in build_review_items(matcher)} == original


def test_已自动确认的小额跨期仍保留跨期事实():
    bank = 记录([("2025-12-31", 3000)])
    journal = 记录([("2026-01-01", 3000)])
    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()
    item = build_review_items(matcher)[0]
    assert item["核对方式"] == "自动确认"
    assert item["跨期"] is True
    assert item["差异金额"] == Decimal("0")


def test_同侧业务链部分已选中时剩余原行不重复归属():
    matcher = Matcher(记录([("2026-01-02", -3000), ("2026-01-03", 3000)]),
                      记录([("2026-01-02", -3000)]), MatcherConfig())
    matcher.run()
    matcher.business_events = [BusinessEvent("E1", "bank", "付款退回", (0, 1), 0, "付款＋退回", "同一业务")]
    items = build_review_items(matcher)
    assert len(items) == 2
    assert sorted(index for item in items for index in item["银行索引"]) == [0, 1]
    assert sum(bool(item["匹配ID"]) for item in items) == 1
    assert all(item["程序结论"] != "同侧业务链" for item in items)


def test_总体真实范围限制覆盖单边且空数据正常返回():
    control = OverallControlResult(scope_limited=True, reasons=("账户身份未验证",))
    item = build_review_items(未匹配([("2026-01-02", 10)], control=control))[0]
    assert item["风险等级"] == "范围未知"
    assert "账户身份未验证" in item["判断依据"]
    assert build_review_items(未匹配([])) == []