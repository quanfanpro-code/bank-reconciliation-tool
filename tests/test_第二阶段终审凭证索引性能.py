"""完整凭证强索引必须先缩小候选池，再做逐组业务证据核验。"""
from decimal import Decimal

import pandas as pd

import matcher as matcher_module
from data_structures import MatcherConfig
from matcher import Matcher
from precision_engine import PrecisionEngine
from 业务分组 import complete_groups


def _记录(金额, 序号, *, 凭证=""):
    return {
        "date": pd.Timestamp("2026-08-10"),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": "设备采购款",
        "aux_text_fields": {
            "摘要": "设备采购款",
            "对方户名": "甲公司",
            "备注": f"第{序号:04d}行",
        },
        "voucher_word": "记" if 凭证 else "",
        "voucher_no": 凭证,
        "original_idx": 序号 + 1,
        "original_file_row": 序号 + 2,
    }


def test_大量同主体同摘要时凭证业务证据核验次数受候选上限约束(monkeypatch):
    银行 = pd.DataFrame([_记录(-200000, 序号) for 序号 in range(200)])
    日记账 = pd.DataFrame([
        _记录(-1, 序号, 凭证="001")
        for 序号 in range(100)
    ])
    配置 = MatcherConfig(max_candidates=30)
    匹配器 = Matcher(银行, 日记账, 配置, logger=lambda _: None)
    凭证组 = complete_groups(
        匹配器._business_rows["journal"],
        配置.dfs_date_window,
        "journal",
    )
    原函数 = matcher_module.business_evidence
    调用数 = 0

    def 计数(银行画像, 日记账画像):
        nonlocal 调用数
        调用数 += 1
        return 原函数(银行画像, 日记账画像)

    monkeypatch.setattr(matcher_module, "business_evidence", 计数)
    匹配器._add_atomic_journal_voucher_candidates(凭证组)

    assert 调用数 <= 配置.max_candidates * 4
    assert len(匹配器.candidates) == 配置.max_candidates
    统计 = 匹配器.run_parameters["candidate_search"]["atomic_voucher"]
    assert 统计["examined"] >= 200
    assert 统计["retained"] == 配置.max_candidates
    assert 统计["truncated_source_rows"] == 1
