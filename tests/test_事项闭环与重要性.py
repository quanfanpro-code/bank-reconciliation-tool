"""核查本次真实缺陷：业务证据、阈值分离和完整事项阅读。"""
from decimal import Decimal

import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from matching_policy import route_candidate
from tests.test_流水对方名称与待办 import _记录


def _匹配器(bank, journal, **settings):
    return Matcher(pd.DataFrame(bank), pd.DataFrame(journal),
                   MatcherConfig(**settings), logger=lambda _: None)


@pytest.mark.parametrize("party,summary", [
    ("黄伟", "2026.5.25黄伟报销差旅费"),
    ("广州药大麦网络科技有限公司", "2026.1.4付广州药大麦网络科技有限公司往来款"),
])
def test_电汇支付方式不构成用途冲突且明确对方可确认(party, summary):
    m = _匹配器([_记录(-120000, "电汇", {"对方账户名": party})],
                 [_记录(-120000, summary, {"业务说明": summary})])
    m.match_exact_1to1()
    m._commit_selected_candidates()
    c = m.selected_candidates[0]
    assert c.evidence["business_strength"] >= 2
    assert not c.evidence["business_conflicts"]
    assert c.processing_status is ProcessingStatus.AUTO_CONFIRMED


def test_小额和改变临界值均不能证明无业务依据的对应():
    statuses = []
    for threshold in (Decimal("0"), Decimal("5000")):
        m = _匹配器([_记录(100, "收款", {})], [_记录(100, "收款", {})],
                     clearly_trivial_threshold=threshold)
        m.match_exact_1to1()
        m._commit_selected_candidates()
        statuses.append(m.selected_candidates[0].processing_status)
    assert statuses == [ProcessingStatus.FLAGGED, ProcessingStatus.FLAGGED]


def test_金额证据评分不随审计临界值改变():
    scores = []
    for threshold in (Decimal("1"), Decimal("5000")):
        m = _匹配器([_记录(100, "设备款", {"对方户名": "甲公司"})],
                     [_记录(90, "设备款", {"对方户名": "甲公司"})],
                     clearly_trivial_threshold=threshold)
        c = m._add_candidate((0,), (0,), "test", "test")
        scores.append(c.scores)
    assert scores[0] == scores[1]


def test_完整凭证候选不受重要性金额窗改变():
    compositions = []
    for materiality in (Decimal("10"), Decimal("100000")):
        bank = [_记录(-1050, "设备结算", {"对方户名": "甲公司"})]
        journal = [_记录(-400, "设备采购", {"对方户名": "甲公司"}, 原表行=2),
                   _记录(-600, "设备采购", {"对方户名": "甲公司"}, 原表行=3)]
        for row in journal:
            row.update(voucher_no="001", voucher_word="记")
        m = _匹配器(bank, journal, performance_materiality=materiality,
                     clearly_trivial_threshold=Decimal("1"))
        m._add_atomic_journal_voucher_candidates([(("凭证", "2026-01", "记", "001", -1), (0, 1))])
        compositions.append({(c.bank_idxs, c.journal_idxs) for c in m.candidates})
    assert compositions == [{((0,), (0, 1))}, {((0,), (0, 1))}]


def test_明确不同对方和真实用途冲突不能自动确认():
    m = _匹配器([_记录(-100, "设备采购", {"对方户名": "甲公司"})],
                 [_记录(-100, "员工工资", {"对方户名": "乙公司"})])
    m.match_exact_1to1()
    m._commit_selected_candidates()
    c = m.selected_candidates[0]
    assert c.evidence["business_conflicts"]
    assert c.processing_status is ProcessingStatus.FLAGGED


def test_同日不同员工报销不能把两笔银行强配成一笔账制造假差额():
    banks = [_记录(-2833.59, "电汇", {"对方账户名": "黄伟", "附言": "报销款"}, 原表行=2),
             _记录(-390.39, "电汇", {"对方账户名": "梁馨玥", "附言": "报销款"}, 原表行=3)]
    journals = [_记录(-2833.59, "2026.5.25黄伟报销差旅费", {"业务说明": "2026.5.25黄伟报销差旅费"}, 原表行=2),
                _记录(-390.39, "2026.5.25梁馨玥报销差旅费", {"业务说明": "2026.5.25梁馨玥报销差旅费"}, 原表行=3)]
    m = _匹配器(banks, journals)
    m._add_candidate((0, 1), (0,), "business_group", "业务组", resolves_full_group=True,
                     complete_business_group=True, represents_full_observed_group=True)
    m.match_exact_1to1()
    from matcher import select_non_conflicting_candidates
    selected = select_non_conflicting_candidates(m.candidates)
    assert all(c.metrics.total_diff_li == 0 for c in selected)
    assert {(c.bank_idxs, c.journal_idxs) for c in selected} == {((0,), (0,)), ((1,), (1,))}


def test_对方科目和开户行不能冒充交易对方():
    from 业务事件 import business_identity
    row = pd.Series(_记录(100, "设备款", {"对方科目名称": "应收账款", "对方行名": "中国银行"}))
    assert business_identity(row)[1] == ""


@pytest.mark.parametrize("materiality,trivial", [("NaN", "1"), ("Infinity", "1"), ("-1", "0"), ("100", "101")])
def test_不合法重要性设置不能进入匹配(materiality, trivial):
    with pytest.raises(ValueError):
        MatcherConfig(performance_materiality=Decimal(materiality), clearly_trivial_threshold=Decimal(trivial))


@pytest.mark.parametrize("replacement,expected_manual,original_day", [("付甲公司设备款", False, "2026-01-09"), ("付乙公司设备款", True, "2026-01-09"), ("付甲公司设备款", True, "2026-04-09")])
def test_有原凭证引用且业务文字相同的同日冲销重记归为闭合更正(replacement, expected_manual, original_day):
    from 复核事项 import build_review_items
    rows = [_记录(-269775, "付甲公司设备款", {}, 日期=original_day, 原表行=2),
            _记录(269775, f"[冲销{original_day.replace('-', '.')} 记-0075号凭证]付甲公司设备款", {}, 日期="2026-03-17", 原表行=3),
            _记录(-269775, replacement, {}, 日期="2026-03-17", 原表行=4)]
    for r, voucher in zip(rows, ("75", "153", "154")):
        r.update(voucher_no=voucher, voucher_word="记")
    m = _匹配器([_记录(-269775, "付甲公司设备款", {}, 日期=original_day, 原表行=2)], rows)
    m.run()
    items = build_review_items(m)
    correction = [i for i in items if set(i["日记账索引"]) & {1, 2}]
    if expected_manual:
        assert any(i["核对方式"].startswith("人工") for i in correction)
    else:
        assert len(correction) == 1
        assert correction[0]["核对方式"] == "留存备查"
        assert correction[0]["日记账索引"] == (1, 2)
        assert correction[0]["净额"] == 0
        assert "冲销重记" in correction[0]["判断依据"]


def test_报告提供完整事项与当日证据且技术表隐藏(tmp_path):
    from tests.test_人工结果计算 import _报告
    from 直读XML import 读簿

    path = _报告(tmp_path, [150, 450, 500], [100, 200, 300, 500, 700],
                 [((0, 1), (0, 1, 2)), ((2,), (3,))])
    from contextlib import nullcontext
    with nullcontext(读簿(path, 读样式=False)) as book:
        assert "事项证据" in book.表名列表
        assert "每日差异证据" in book.表名列表
        overview = {r[0]: r[1:] for r in list(book["核对总览"].行值())[1:]}
        assert overview["实际执行重要性水平"][0] == 50
        assert overview["明显微小错报临界值"][0] == 1
        assert overview["初始真正单边记录数"][0] == 1
        pending = list(book["待处理事项"].行值())
        assert "程序归纳与具体问题" in pending[0]
        assert "需要你回答" in pending[1][pending[0].index("程序归纳与具体问题")]
        text = str(list(book["事项证据"].行值()))
        assert "150.00" in text and "450.00" in text and "600.00" in text
        assert all(book[name].状态 == "hidden" for name in ("复核事项索引", "核对明细", "人工全查", "运行参数"))
        daily = book["每日统计"]
        assert list(daily.行值())[0][19] == "需查看"
        assert "每日差异证据" in str(daily.值(2, 20))
