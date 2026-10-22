"""复现353项中实际发现的用途误判及完整凭证过度拦截。"""
import pytest
from data_structures import ProcessingStatus
from tests.test_事项闭环与重要性 import _匹配器
from tests.test_流水对方名称与待办 import _记录


@pytest.mark.parametrize("bank_text,party,journal_text,extra", [
    ("税费等", "税费社保", "2026.2.26缴纳1月增值税及附加税", {}),
    ("税费等", "税费社保", "支付1月社保", {}),
    ("", "苟小敏", "2026.7.23苟小敏费用报销", {}),
    ("电汇", "苟小敏", "2026.4.2收回苟小敏备用金", {"附言": "备用金退回"}),
    ("电汇", "重庆药大麦医药科技有限公司", "2026.2.26往来转账", {"附言": "转款"}),
])
def test_具体业务证据不被开户行和科目名称污染(bank_text, party, journal_text, extra):
    bank = _记录(-2500000, bank_text, {"对方账户名": party, "对方行名": "中国银行", **extra})
    journal = _记录(-2500000, journal_text, {"业务说明": journal_text, "对方科目名称": "银行存款" if "转账" in journal_text else "应付职工薪酬"})
    m = _匹配器([bank], [journal])
    m.match_exact_1to1()
    m._commit_selected_candidates()
    c = m.selected_candidates[0]
    assert not c.evidence["business_conflicts"]
    assert c.processing_status is ProcessingStatus.AUTO_CONFIRMED


def test_完整凭证可以包含多笔有独立确切证据的支付():
    banks = [_记录(-22849.46, "税费等", {"对方账户名": "税费社保"}, 原表行=2),
             _记录(-10861.12, "税费等", {"对方账户名": "税费社保"}, 原表行=3)]
    journals = [_记录(-22849.46, "支付1月社保", {"业务说明": "支付1月社保"}, 原表行=2),
                _记录(-10861.12, "支付1月社保", {"业务说明": "支付1月社保"}, 原表行=3)]
    m = _匹配器(banks, journals)
    m._atomic_groups["journal"] = [frozenset((0, 1))]
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert len(m.selected_candidates) == 2
    assert all(c.processing_status is ProcessingStatus.AUTO_CONFIRMED for c in m.selected_candidates)


def test_明确不同用途仍然是人工疑点():
    m = _匹配器([_记录(-120000, "设备采购", {"对方账户名": "苟小敏"})],
                [_记录(-120000, "苟小敏报销差旅费", {"业务说明": "苟小敏报销差旅费"})])
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert m.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED


def test_同日同对方重复金额可整组核对无需猜每笔对应():
    party = "广州药大麦网络科技有限公司"
    banks = [_记录(-100000, "", {"对方账户名": party}, 原表行=n) for n in (2, 3)]
    journals = [_记录(-100000, f"付{party}{purpose}", {"业务说明": f"付{party}{purpose}"}, 原表行=n)
                for n, purpose in ((2, "管理费"), (3, "推广费"))]
    m = _匹配器(banks, journals)
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert c.bank_idxs == c.journal_idxs == (0, 1)
    assert c.processing_status is ProcessingStatus.GROUP_RECONCILED


def test_同日逐笔确证不会被跨日带差额的同用途拼组否定():
    banks = [_记录(-8000, "电汇", {"附言": "工资发放"}, 原表行=2),
             _记录(-9000, "电汇", {"附言": "工资发放"}, 原表行=3)]
    journals = [_记录(-8000, "发放1月工资", {}, 原表行=2),
                _记录(-18000, "工资", {}, 日期="2026-01-12", 原表行=3)]
    m = _匹配器(banks, journals)
    m.match_exact_1to1()
    m._add_candidate((0, 1), (1,), "business_group", "业务完整组",
                     complete_business_group=True, represents_full_observed_group=True)
    m._commit_selected_candidates()
    c = next(c for c in m.candidates if c.bank_idxs == c.journal_idxs == (0,))
    assert not c.is_ambiguous


def test_银行未载工资用途时完整凭证整理为一项但不自动确认():
    banks = [_记录(-10000, "", {"对方账户名": "甲员工"}, 原表行=2),
             _记录(-8000, "", {"对方账户名": "乙员工"}, 原表行=3)]
    journals = [_记录(-10000, "发放6月工资", {}, 原表行=2), _记录(-8000, "发放6月工资", {}, 原表行=3)]
    m = _匹配器(banks, journals)
    m._atomic_groups["journal"] = [frozenset((0, 1))]
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert len(m.selected_candidates) == 1
    c = m.selected_candidates[0]
    assert c.bank_idxs == c.journal_idxs == (0, 1)
    assert c.processing_status is ProcessingStatus.FLAGGED
    assert not c.evidence.get("resolves_full_group")


def test_唯一同日明确银行冲正自行闭合但保留原始收支():
    from 业务事件 import detect_same_side_events
    from 复核事项 import build_review_items
    banks = [_记录(-71440, "电汇", {"对方账户名": "本草数链（深圳）科技有限公司"}, 原表行=2),
             _记录(71440, "冲正-电汇", {"对方账户名": "本草数链（深圳）科技有限公司", "附言": "收款行拒绝入账"}, 原表行=3)]
    m = _匹配器(banks, [_记录(1, "其他业务", {})])
    m.business_events = detect_same_side_events(m.bank, "bank")[0]
    item = next(i for i in build_review_items(m) if i["银行索引"])
    assert item["程序结论"] == "银行冲正已闭合"
    assert item["核对方式"] == "留存备查"
    assert item["银行收入"] == item["银行支出"] == 71440


def test_银行附言直接保留在判断证据中():
    from 复核报表 import _brief
    assert "报销款" in _brief(_记录(-100, "电汇", {"附言": "报销款"}))


@pytest.mark.parametrize("purpose", ["社保扣款", "设备采购"])
def test_税费社保汇总户名不能覆盖银行明示的不同用途(purpose):
    m = _匹配器([_记录(-43210, purpose, {"对方账户名": "税费社保"})],
                [_记录(-43210, "缴纳增值税", {"业务说明": "缴纳增值税"})])
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert m.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED


def test_余额碰巧相同不能覆盖明确的不同用途():
    b = _记录(123456, "电费", {"对方账户名": "甲单位"})
    j = _记录(123456, "往来转款", {"对方科目名称": "银行存款"})
    b["balance"] = j["balance"] = 345678
    m = _匹配器([b], [j])
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert m.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED


def test_共同订单只能命中部分凭证时不能绕过完整范围():
    m = _匹配器([_记录(-17000, "采购款", {"订单号": "ORD-A"})],
                [_记录(-17000, "采购款", {"订单号": "ORD-A"}, 原表行=2),
                 _记录(-23000, "采购款", {"订单号": "ORD-B"}, 原表行=3)])
    m._atomic_groups["journal"] = [frozenset((0, 1))]
    m.match_exact_1to1()
    m._commit_selected_candidates()
    assert not m.selected_candidates


def test_原付款不唯一时银行冲正仍须保留人工疑点():
    from 业务事件 import detect_same_side_events
    from 复核事项 import build_review_items
    banks = [_记录(-12345, "电汇", {"对方账户名": "甲公司"}, 原表行=n) for n in (2, 3)]
    banks.append(_记录(12345, "冲正-电汇", {"对方账户名": "甲公司"}, 原表行=4))
    m = _匹配器(banks, [_记录(1, "其他", {})])
    m.business_events = detect_same_side_events(m.bank, "bank")[0]
    assert not any(i["程序结论"] == "银行冲正已闭合" for i in build_review_items(m))
