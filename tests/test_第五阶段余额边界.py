"""独立复核余额乱序及局部断档的结构性限制。"""

from decimal import Decimal
from itertools import permutations

import pandas as pd
import pytest

from balance import BalanceRecalculator, build_overall_controls, check_balance_continuity


def _表(行):
    表 = pd.DataFrame(行, columns=["date", "amount", "balance"])
    表["date"] = pd.to_datetime(表["date"])
    for 列 in ("amount", "balance"):
        表[列] = 表[列].map(lambda 值: None if 值 is None else Decimal(str(值)))
    表["original_file_row"] = range(2, len(表) + 2)
    return 表


@pytest.mark.parametrize("顺序", list(permutations(range(3))))
@pytest.mark.parametrize("位置", ["首日", "中间日", "末日"])
@pytest.mark.parametrize("混合收支", [False, True])
def test_完整余额链不因同日首末行乱序而误报(顺序, 位置, 混合收支):
    当日 = [(100, 5100), (-250, 4850), (200, 5050)] if 混合收支 else [(100, 5100), (300, 5400), (200, 5600)]
    行 = [("2026-01-06", *当日[i]) for i in 顺序]
    if 位置 != "首日":
        行.insert(0, ("2026-01-05", 5000, 5000))
    if 位置 != "末日":
        行.append(("2026-01-07", 50, 当日[-1][1] + 50))
    表 = _表(行)
    原表 = 表.copy(deep=True)
    assert check_balance_continuity(表) == []
    控制 = build_overall_controls(表, 表.copy())
    assert 控制.scope_limited is False
    assert 控制.bank_balance_diff == 0
    assert BalanceRecalculator.extract_initial_balance(表) == (5000 if 位置 == "首日" else 0)
    pd.testing.assert_frame_equal(表, 原表)


def test_局部断档不能豁免另一侧余额点不足():
    银行 = _表([(f"2026-01-{日:02}", 金额, 余额) for 日, 金额, 余额 in
                [(5, 5000, 5000), (6, 100, 4100), (7, 50, 4150), (8, 60, 4210), (9, 70, 4280)]])
    账 = 银行.copy()
    账["balance"] = [None, None, None, None, Decimal("5280")]
    控制 = build_overall_controls(银行, 账)
    assert 控制.scope_limited is True
    assert any("仅有1个有效余额点" in 原因 for 原因 in 控制.reasons)


def test_完整重排不得掩盖真实缺行():
    表 = _表([("2026-01-05", 5000, 5000), ("2026-01-06", 300, 5300),
              ("2026-01-06", 100, 5000), ("2026-01-06", 200, 5500),
              ("2026-01-07", 50, 5550)])
    异常 = check_balance_continuity(表)
    assert 异常
    assert any(项["差额"] == Decimal("100") for 项 in 异常)

@pytest.mark.parametrize("顺序", [(0, 1), (1, 0)])
def test_末日收付相抵的闭环沿用已证明的前日余额(顺序):
    当日 = [(100, 5100), (-100, 5000)]
    表 = _表([("2026-01-05", 5000, 5000)] + [("2026-01-06", *当日[i]) for i in 顺序])
    assert check_balance_continuity(表) == []
    控制 = build_overall_controls(表, 表.copy())
    assert 控制.bank_ending_balance == Decimal("5000")
    assert 控制.bank_balance_diff == 0
    assert 控制.scope_limited is False
