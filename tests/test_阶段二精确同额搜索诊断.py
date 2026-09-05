from decimal import Decimal
import time

import pandas as pd
import pytest

from data_structures import MatcherConfig
import matcher as 匹配模块
from matcher import Matcher
from precision_engine import PrecisionEngine


def _数据(金额列表, 来源):
    return pd.DataFrame([
        {
            "date": pd.Timestamp("2026-08-31"),
            "amount": Decimal(str(金额)),
            "amount_decimal": PrecisionEngine.to_integer_li(金额),
            "summary": "",
            "aux_text_fields": {},
            "balance": None,
            "voucher_no": "",
            "source": 来源,
            "original_idx": 序号,
            "original_file_row": 序号 + 2,
        }
        for 序号, 金额 in enumerate(金额列表)
    ])


@pytest.mark.parametrize("同额数量", [30, 31, 100])
def test_全是精确同额且没有可补搜拆分时不得误报搜索未穷尽(同额数量):
    匹配器 = Matcher(
        _数据([100], "bank"),
        _数据([100] * 同额数量, "journal"),
        MatcherConfig(allow_greedy_fallback=False),
    )

    匹配器.match_dfs_combinations()

    搜索 = 匹配器.run_parameters["combination_search"]
    assert 搜索["generic_source_rows"] == 1
    assert 搜索["fully_searched_source_rows"] == 1
    assert 搜索["depth_limited_source_rows"] == 0
    assert 搜索["budget_exhausted_source_rows"] == 0


def test_全局预算必须覆盖候选窗口业务证据扫描(monkeypatch):
    调用次数 = 0
    原函数 = 匹配模块.business_evidence

    def 计数(*args, **kwargs):
        nonlocal 调用次数
        调用次数 += 1
        return 原函数(*args, **kwargs)

    monkeypatch.setattr(匹配模块, "business_evidence", 计数)
    匹配器 = Matcher(
        _数据([100] * 200, "bank"),
        _数据([100] * 200, "journal"),
        MatcherConfig(
            allow_greedy_fallback=False,
            combination_global_time_limit_seconds=0,
        ),
    )

    开始 = time.perf_counter()
    匹配器.match_dfs_combinations()
    耗时 = time.perf_counter() - 开始

    搜索 = 匹配器.run_parameters["combination_search"]
    assert 调用次数 == 0
    assert 搜索["global_timeout_unprocessed_source_rows"] == 400
    assert 搜索["budget_exhausted_source_rows"] == 400
    assert 耗时 < 2
