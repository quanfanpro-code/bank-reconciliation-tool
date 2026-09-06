"""月度、每日余额采用可证明的原始余额链尾，不能用重算抹平断档。"""

from decimal import Decimal
from itertools import permutations

import pandas as pd
import pytest

from balance import daily_balance_points, check_balance_continuity


def _流水(行):
    表 = pd.DataFrame(行, columns=["date", "amount", "balance"])
    表["date"] = pd.to_datetime(表["date"], errors="coerce")
    for 列 in ("amount", "balance"):
        表[列] = 表[列].map(lambda 值: None if 值 is None else Decimal(str(值)))
    表["original_file_row"] = range(2, len(表) + 2)
    return 表


@pytest.mark.parametrize("顺序", list(permutations(range(3))))
def test_同日乱序仍取完整余额链尾而非文件末行(顺序):
    当日 = [(100, 5100), (300, 5400), (200, 5600)]
    表 = _流水([("2026-01-31", *当日[i]) for i in 顺序])
    原表 = 表.copy(deep=True)

    assert daily_balance_points(表) == {pd.Timestamp("2026-01-31"): Decimal("5600")}
    pd.testing.assert_frame_equal(表, 原表)


@pytest.mark.parametrize("顺序", [(0, 1), (1, 0)])
def test_末日净零闭环采用前日可信余额(顺序):
    当日 = [(100, 5100), (-100, 5000)]
    表 = _流水([("2026-01-30", 5000, 5000)] + [("2026-01-31", *当日[i]) for i in 顺序])

    assert daily_balance_points(表) == {
        pd.Timestamp("2026-01-30"): Decimal("5000"),
        pd.Timestamp("2026-01-31"): Decimal("5000"),
    }


def test_净零闭环没有可靠前日锚点时不猜末余额():
    表 = _流水([("2026-01-31", 100, 5100), ("2026-01-31", -100, 5000)])

    assert daily_balance_points(表) == {pd.Timestamp("2026-01-31"): None}


def test_无余额交易日明确缺失并可为后续闭环累计已知净额():
    表 = _流水([
        ("2026-01-28", 5000, 5000),
        ("2026-01-29", 100, None),
        ("2026-01-31", -200, 5100),
        ("2026-01-31", 200, 5300),
    ])

    assert daily_balance_points(表) == {
        pd.Timestamp("2026-01-28"): Decimal("5000"),
        pd.Timestamp("2026-01-29"): None,
        pd.Timestamp("2026-01-31"): Decimal("5100"),
    }


def test_真实日界断档保留原余额和异常而非按期初重算():
    表 = _流水([
        ("2026-01-30", 5000, 5000),
        ("2026-01-31", 100, 4100),
        ("2026-01-31", 300, 4400),
    ])
    原异常 = check_balance_continuity(表)

    assert daily_balance_points(表)[pd.Timestamp("2026-01-31")] == Decimal("4400")
    assert check_balance_continuity(表) == 原异常
    assert any(项["差额"] == Decimal("1000") for 项 in 原异常)


def test_日内余额链不完整时不选最后一行冒充可信期末():
    表 = _流水([
        ("2026-01-30", 5000, 5000),
        ("2026-01-31", 100, 5100),
        ("2026-01-31", 200, 5600),
    ])

    assert daily_balance_points(表)[pd.Timestamp("2026-01-31")] is None
    assert check_balance_continuity(表)


def test_末笔缺余额不结转前一笔并不污染后续净零锚点():
    表 = _流水([
        ("2026-01-29", 5000, 5000),
        ("2026-01-30", 100, 5100),
        ("2026-01-30", 200, None),
        ("2026-01-31", -100, 5200),
        ("2026-01-31", 100, 5300),
    ])

    点 = daily_balance_points(表)
    assert 点[pd.Timestamp("2026-01-30")] is None
    assert 点[pd.Timestamp("2026-01-31")] is None


def test_缺失余额列不造零且有效零余额不会丢失():
    表 = _流水([("2026-01-31", -5000, 0)])

    assert daily_balance_points(表) == {pd.Timestamp("2026-01-31"): Decimal("0")}
    assert daily_balance_points(表.drop(columns="balance")) == {pd.Timestamp("2026-01-31"): None}
    assert daily_balance_points(pd.DataFrame()) == {}


def test_日期归一到日且不为无交易日制造点位():
    表 = _流水([("2026-01-29 10:00", 5000, 5000), ("2026-01-31 12:00", 100, 5100)])
    表.loc[2] = [pd.NaT, Decimal("9000"), Decimal("99999"), 4]

    assert daily_balance_points(表) == {
        pd.Timestamp("2026-01-29"): Decimal("5000"),
        pd.Timestamp("2026-01-31"): Decimal("5100"),
    }
