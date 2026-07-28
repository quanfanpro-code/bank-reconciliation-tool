from decimal import Decimal

import pandas as pd
import pytest

import data_loader as data_loader_module
from balance import BalanceRecalculator
from data_loader import DataLoader, ParseErrorCollector


def _single_amount_result(source_type: str, direction: str) -> Decimal:
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-02"],
            "金额": [100],
            "方向": [direction],
            "摘要": ["测试"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "summary": "摘要",
        "mode": "single_amount_with_direction",
    }
    result = DataLoader().standardize_data(frame, mapping, source_type)
    return result.iloc[0]["amount"]


@pytest.mark.parametrize(
    ("source_type", "direction", "expected"),
    [
        ("bank", "贷", Decimal("100.00")),
        ("bank", "借", Decimal("-100.00")),
        ("bank", "CR", Decimal("100.00")),
        ("bank", "DR", Decimal("-100.00")),
        ("journal", "借", Decimal("100.00")),
        ("journal", "贷", Decimal("-100.00")),
        ("journal", "DR", Decimal("100.00")),
        ("journal", "CR", Decimal("-100.00")),
    ],
)
def test_单列金额方向必须区分银行和日记账(source_type, direction, expected):
    assert _single_amount_result(source_type, direction) == expected


def test_借贷分列继续按银行贷增借减和日记账借增贷减():
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-01", "2026-01-02"],
            "借方": [100, 0],
            "贷方": [0, 40],
            "摘要": ["借方一笔", "贷方一笔"],
        }
    )
    mapping = {
        "date": "日期",
        "debit": "借方",
        "credit": "贷方",
        "summary": "摘要",
        "mode": "debit_credit",
    }

    bank = DataLoader().standardize_data(frame, mapping, "bank")
    journal = DataLoader().standardize_data(frame, mapping, "journal")

    assert dict(zip(bank["summary"], bank["amount"])) == {
        "借方一笔": Decimal("-100.00"),
        "贷方一笔": Decimal("40.00"),
    }
    assert dict(zip(journal["summary"], journal["amount"])) == {
        "借方一笔": Decimal("100.00"),
        "贷方一笔": Decimal("-40.00"),
    }


@pytest.mark.parametrize(
    ("source_type", "value", "expected"),
    [
        ("bank", "100CR", Decimal("100.00")),
        ("bank", "100DR", Decimal("-100.00")),
        ("journal", "100CR", Decimal("-100.00")),
        ("journal", "100DR", Decimal("100.00")),
    ],
)
def test_带CR_DR后缀的金额也必须区分数据来源(source_type, value, expected):
    frame = pd.DataFrame(
        {"日期": ["2026-01-01"], "金额": [value], "摘要": ["测试"]}
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "mode": "signed_amount",
    }

    result = DataLoader().standardize_data(frame, mapping, source_type)

    assert result.iloc[0]["amount"] == expected


def test_未知方向单独记为方向解析失败而不是金额失败():
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-01", "2026-01-02"],
            "金额": [50, 100],
            "方向": ["贷", "未知"],
            "摘要": ["正常", "测试"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "summary": "摘要",
        "mode": "single_amount_with_direction",
    }

    result = loader.standardize_data(frame, mapping, "bank")
    summary = collector.get_summary()

    assert len(result) == 1
    assert summary["方向解析失败"] == 1
    assert summary["金额解析失败"] == 0


def test_单列非法金额必须进入金额解析异常():
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-01", "2026-01-02"],
            "金额": [50, "不是金额"],
            "摘要": ["正常", "测试"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "mode": "signed_amount",
    }

    result = loader.standardize_data(frame, mapping, "bank")

    assert len(result) == 1
    assert collector.get_summary()["金额解析失败"] == 1


def test_标准化结果保留多个带字段名的辅助文字():
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-01"],
            "金额": [100],
            "摘要": ["货款"],
            "业务说明": ["销售回款"],
            "对方户名": ["甲公司"],
            "凭证号": ["记-001"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "voucher": "凭证号",
        "auxiliary_text_columns": ["摘要", "业务说明", "对方户名"],
        "mode": "signed_amount",
    }

    result = DataLoader().standardize_data(frame, mapping, "bank")

    assert result.iloc[0]["aux_text_fields"] == {
        "摘要": "货款",
        "业务说明": "销售回款",
        "对方户名": "甲公司",
    }
    assert "凭证号" not in result.iloc[0]["aux_text_fields"]


def test_日记账单列方向必须参与期初余额推算():
    frame = pd.DataFrame(
        {
            "日期": ["2026-01-01"],
            "金额": [100],
            "方向": ["贷"],
            "余额": [900],
            "摘要": ["付款"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "balance": "余额",
        "summary": "摘要",
        "mode": "single_amount_with_direction",
    }

    initial = BalanceRecalculator.extract_initial_balance(
        frame,
        mapping,
        source_type="journal",
    )

    assert initial == Decimal("1000.00")


def test_方向解析函数对未知数据来源拒绝猜测():
    with pytest.raises(ValueError, match="数据来源"):
        data_loader_module.direction_sign("借", "unknown")
