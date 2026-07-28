"""
核心功能自检 — 单元测试
从原 self_test() 函数拆分而来
"""
from decimal import Decimal
from datetime import datetime

import pandas as pd
import pytest

from utils import parse_date, clean_amount


# ==========================================
# 金额解析测试
# ==========================================


def test_金额解析_千分位格式():
    assert clean_amount("1,234.56") == Decimal("1234.56")


def test_金额解析_非法逗号格式返回None():
    assert clean_amount("1,2,3") is None


def test_金额解析_负数格式():
    assert clean_amount("-100.00") == Decimal("-100.00")


def test_金额解析_人民币符号():
    assert clean_amount("¥2,000.50") == Decimal("2000.50")


def test_金额解析_括号负数():
    assert clean_amount("(500.00)") == Decimal("-500.00")


def test_金额解析_CR后缀():
    assert clean_amount("100CR") == Decimal("-100.00")


def test_金额解析_DR后缀():
    assert clean_amount("100DR") == Decimal("100.00")


def test_金额解析_贷后缀():
    assert clean_amount("100贷") == Decimal("-100.00")


def test_金额解析_借后缀():
    assert clean_amount("100借") == Decimal("100.00")


def test_金额解析_空字符串():
    assert clean_amount("") == Decimal("0.00")


def test_金额解析_None值():
    assert clean_amount(None) == Decimal("0.00")


def test_金额解析_浮点数四舍五入():
    assert clean_amount(123.456) == Decimal("123.46")


def test_金额解析_全角数字():
    assert clean_amount("１２３４.５６") == Decimal("1234.56")


def test_金额解析_无效字符串返回None():
    assert clean_amount("invalid") is None


# ==========================================
# 日期解析测试
# ==========================================


def test_日期解析_YYYY_MM_DD格式():
    expected = pd.Timestamp("2024-03-15")
    assert parse_date("2024-03-15") == expected


def test_日期解析_斜杠格式():
    expected = pd.Timestamp("2024-03-15")
    assert parse_date("2024/03/15") == expected


def test_日期解析_紧凑格式():
    expected = pd.Timestamp("2024-03-15")
    assert parse_date("20240315") == expected


def test_日期解析_datetime对象():
    expected = pd.Timestamp("2024-03-15")
    result = parse_date(datetime(2024, 3, 15, 14, 30, 0))
    assert result == expected


def test_日期解析_Excel序列号():
    result = parse_date(45336)
    assert result is not None and result.year == 2024


def test_日期解析_None返回None():
    assert parse_date(None) is None


def test_日期解析_空字符串返回None():
    assert parse_date("") is None
