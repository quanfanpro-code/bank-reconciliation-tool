"""通用组合预处理必须先限候选池，再做昂贵业务证据核验。"""

from decimal import Decimal
import time

import pandas as pd

from data_structures import MatcherConfig
import matcher as 匹配模块
from matcher import Matcher
from precision_engine import PrecisionEngine


def _数据(金额列表, 来源):
    return pd.DataFrame(
        [
            {
                "date": pd.Timestamp("2026-08-31"),
                "amount": Decimal(str(金额)),
                "amount_decimal": PrecisionEngine.to_integer_li(金额),
                "summary": "普通货款",
                "aux_text_fields": {
                    "摘要": "普通货款",
                    "对方户名": "甲公司",
                },
                "balance": None,
                "voucher_word": "",
                "voucher_no": "",
                "source": 来源,
                "original_idx": 序号,
                "original_file_row": 序号 + 2,
            }
            for 序号, 金额 in enumerate(金额列表)
        ]
    )


def _准备单向通用组合(匹配器):
    匹配器._combination_covered = {"bank": set(), "journal": set()}
    匹配器._reset_candidate_search_stage("generic_combination")
    匹配器._combination_source_sets = {
        "generic": set(),
        "fully_searched": set(),
        "truncated": set(),
        "depth_limited": set(),
        "node_budget_exhausted": set(),
        "task_timeout": set(),
        "global_timeout_unprocessed": set(),
        "worker_failure": set(),
        "budget_exhausted": set(),
    }
    匹配器._combination_global_deadline = time.monotonic() + 60
    匹配器.run_parameters["combination_search"] = {
        "candidate_limit": min(匹配器.config.max_candidates, 30),
    }


def test_二百乘二百普通货款预处理的业务证据调用受候选上限约束(monkeypatch):
    数量, 上限 = 200, 30
    匹配器 = Matcher(
        _数据([100] * 数量, "bank"),
        _数据([40] * 数量, "journal"),
        MatcherConfig(max_candidates=上限, allow_greedy_fallback=False),
        logger=lambda _: None,
    )
    _准备单向通用组合(匹配器)
    原函数 = 匹配模块.business_evidence
    调用数 = 0

    def 计数(银行画像, 日记账画像):
        nonlocal 调用数
        调用数 += 1
        return 原函数(银行画像, 日记账画像)

    def 跳过后续组合求解(参数):
        return 匹配模块._CombinationTaskResult(
            source_idx=int(参数[0]),
            legacy_result=None,
            candidate_count=0,
            retained_count=0,
            depth_limited=False,
            nodes_visited=0,
        )

    monkeypatch.setattr(匹配模块, "business_evidence", 计数)
    monkeypatch.setattr(
        匹配模块,
        "_process_single_source_with_diagnostics",
        跳过后续组合求解,
    )
    monkeypatch.setattr(匹配器, "_should_use_parallel", lambda _: False)

    匹配器._dfs_one_to_many("bank", "journal")

    assert 调用数 <= 数量 * 上限 * 4
    统计 = 匹配器.run_parameters["candidate_search"]["generic_combination"]
    assert 统计 == {
        "examined": 数量 * 数量,
        "retained": 数量 * 上限,
        "truncated_source_rows": 数量,
    }


def test_候选池限额后仍能找到小型一对多正确拆分():
    匹配器 = Matcher(
        _数据([100], "bank"),
        _数据([40, 60], "journal"),
        MatcherConfig(max_candidates=30, allow_greedy_fallback=False),
        logger=lambda _: None,
    )
    _准备单向通用组合(匹配器)

    匹配器._dfs_one_to_many("bank", "journal")

    assert any(
        候选.bank_idxs == (0,)
        and 候选.journal_idxs == (0, 1)
        and 候选.metrics.total_diff_li == 0
        for 候选 in 匹配器.candidates
    )
    assert 匹配器.run_parameters["candidate_search"]["generic_combination"] == {
        "examined": 2,
        "retained": 2,
        "truncated_source_rows": 0,
    }
