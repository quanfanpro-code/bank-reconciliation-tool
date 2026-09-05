"""完整业务组必须先稳定缩小候选池，并准确披露候选搜索范围。"""

from decimal import Decimal

import pandas as pd

import matcher as matcher_module
from data_structures import MatcherConfig
from matcher import Matcher
from precision_engine import PrecisionEngine


def _记录(
    金额,
    序号,
    *,
    摘要,
    对方,
    业务编号="",
    日期="2026-08-31",
    凭证字="",
    凭证号="",
):
    辅助文字 = {"摘要": 摘要, "对方户名": 对方}
    if 业务编号:
        辅助文字["业务编号"] = 业务编号
    return {
        "date": pd.Timestamp(日期),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 辅助文字,
        "balance": None,
        "voucher_word": 凭证字,
        "voucher_no": 凭证号,
        "original_idx": 序号 + 1,
        "original_file_row": 序号 + 2,
    }


def test_二百乘二百同额业务组的业务证据核验必须受候选上限约束(monkeypatch):
    数量 = 200
    上限 = 7
    银行 = pd.DataFrame([
        _记录(-100, 序号, 摘要=f"项目{序号:04d}服务款", 对方=f"公司{序号:04d}")
        for 序号 in range(数量)
    ])
    日记账 = 银行.copy(deep=True)
    配置 = MatcherConfig(max_candidates=上限)
    匹配器 = Matcher(银行, 日记账, 配置, logger=lambda _: None)

    原函数 = matcher_module.business_evidence
    调用数 = 0

    def 计数(银行画像, 日记账画像):
        nonlocal 调用数
        调用数 += 1
        return 原函数(银行画像, 日记账画像)

    monkeypatch.setattr(matcher_module, "business_evidence", 计数)
    匹配器.match_business_groups()

    assert 调用数 <= 数量 * 上限
    assert 匹配器.run_parameters["candidate_search"]["business_group"] == {
        "examined": 数量 * 数量,
        "retained": 调用数,
        "truncated_source_rows": 数量,
    }


def test_缩小候选池后仍保留共同业务编号的完整组正例():
    银行 = pd.DataFrame([
        _记录(-300, 0, 摘要="项目结算", 对方="甲公司", 业务编号="B-001")
    ])
    日记账 = pd.DataFrame([
        _记录(-100, 0, 摘要="项目结算", 对方="甲公司", 业务编号="B-001"),
        _记录(-200, 1, 摘要="项目结算", 对方="甲公司", 业务编号="B-001"),
    ])
    匹配器 = Matcher(
        银行,
        日记账,
        MatcherConfig(max_candidates=7),
        logger=lambda _: None,
    )

    匹配器.match_business_groups()

    assert any(
        候选.bank_idxs == (0,)
        and 候选.journal_idxs == (0, 1)
        and 候选.evidence.get("shared_business_id")
        for 候选 in 匹配器.candidates
    )
    assert 匹配器.run_parameters["candidate_search"]["business_group"] == {
        "examined": 1,
        "retained": 1,
        "truncated_source_rows": 0,
    }


def test_同业务编号跨期间候选必须先按日期窗保留同日正确组():
    银行 = pd.DataFrame([
        _记录(
            -100,
            0,
            摘要="当前项目结算",
            对方="甲公司",
            业务编号="X-1",
            日期="2026-12-31",
        )
    ])
    日记账 = pd.DataFrame([
        _记录(
            -999,
            0,
            摘要="远端项目结算",
            对方="甲公司",
            业务编号="X-1",
            日期="2026-01-01",
        ),
        _记录(
            -40,
            1,
            摘要="当前项目结算",
            对方="甲公司",
            业务编号="X-1",
            日期="2026-12-31",
        ),
        _记录(
            -60,
            2,
            摘要="当前项目结算",
            对方="甲公司",
            业务编号="X-1",
            日期="2026-12-31",
        ),
    ])
    匹配器 = Matcher(
        银行,
        日记账,
        MatcherConfig(max_candidates=1, dfs_date_window=31),
        logger=lambda _: None,
    )

    匹配器.match_business_groups()

    assert any(
        候选.bank_idxs == (0,)
        and 候选.journal_idxs == (1, 2)
        and 候选.evidence.get("shared_business_id")
        for 候选 in 匹配器.candidates
    )
    assert 匹配器.run_parameters["candidate_search"]["business_group"] == {
        "examined": 2,
        "retained": 1,
        "truncated_source_rows": 1,
    }


def test_同日多个共享业务编号凭证必须优先保留精确总额组(monkeypatch):
    凭证组数 = 20
    候选上限 = 1
    银行 = pd.DataFrame([
        _记录(
            -100,
            0,
            摘要="当前项目结算",
            对方="甲公司",
            业务编号="X-1",
            日期="2026-12-31",
        )
    ])
    日记账记录 = []
    for 组号 in range(凭证组数):
        组总额 = -100 if 组号 == 0 else -(200 + 组号)
        公共参数 = {
            "摘要": f"凭证项目{组号}",
            "对方": "甲公司" if 组号 == 0 else f"公司{组号}",
            "业务编号": "X-1",
            "日期": "2026-12-31",
            "凭证字": "记",
            "凭证号": str(组号 + 1),
        }
        日记账记录.extend([
            _记录(-40, 组号 * 2, **公共参数),
            _记录(组总额 + 40, 组号 * 2 + 1, **公共参数),
        ])
    匹配器 = Matcher(
        银行,
        pd.DataFrame(日记账记录),
        MatcherConfig(max_candidates=候选上限, dfs_date_window=31),
        logger=lambda _: None,
    )

    原函数 = matcher_module.business_evidence
    调用数 = 0

    def 计数(银行画像, 日记账画像):
        nonlocal 调用数
        调用数 += 1
        return 原函数(银行画像, 日记账画像)

    monkeypatch.setattr(matcher_module, "business_evidence", 计数)
    匹配器.match_business_groups()

    assert any(
        候选.bank_idxs == (0,)
        and 候选.journal_idxs == (0, 1)
        and 候选.metrics.total_diff_li == 0
        and 候选.evidence.get("shared_business_id")
        for 候选 in 匹配器.candidates
    )
    assert 调用数 <= 1 + 凭证组数 * 候选上限 * 4
    assert 匹配器.run_parameters["candidate_search"]["business_group"] == {
        "examined": 凭证组数,
        "retained": 1,
        "truncated_source_rows": 1,
    }
    assert 匹配器.run_parameters["candidate_search"]["atomic_voucher"] == {
        "examined": 凭证组数,
        "retained": 1,
        "truncated_source_rows": 凭证组数 - 1,
    }
