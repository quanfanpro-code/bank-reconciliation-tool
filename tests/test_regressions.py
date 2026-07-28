from decimal import Decimal
import warnings

import numpy as np
import pandas as pd

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig
from matcher import Matcher, _randomized_greedy
from validate import validate_config_params


def test_保留原始行号用于回填原始数据():
    df = pd.DataFrame(
        {
            "日期": ["2024-01-01", "", "2024-01-02"],
            "摘要": ["收款", "本期合计", "付款"],
            "凭证": ["001", "", "002"],
            "借方": [100, "", 0],
            "贷方": [0, "", 50],
        }
    )

    loader = DataLoader()
    mapping = {
        "date": "日期",
        "summary": "摘要",
        "voucher": "凭证",
        "debit": "借方",
        "credit": "贷方",
        "balance": None,
        "mode": "debit_credit",
    }

    result = loader.standardize_data(df, mapping, "journal", "auto")

    assert result["original_idx"].tolist() == [1, 3]


def test_日期异常统计会被正确记录():
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    df = pd.DataFrame(
        {
            "日期": ["坏日期", ""],
            "金额": [100, 200],
            "摘要": ["异常日期", "空日期"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "voucher": None,
        "balance": None,
        "mode": "signed_amount",
    }

    result = loader.standardize_data(df, mapping, "bank", "auto")
    summary = collector.get_summary()

    assert result.empty
    assert summary["日期解析失败"] == 1
    assert summary["空日期行"] == 1


def test_原文件行号会叠加跳过行偏移():
    df = pd.DataFrame(
        {
            "日期": ["2024-01-01"],
            "金额": [100],
            "摘要": ["正常记录"],
        }
    )

    loader = DataLoader()
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "voucher": None,
        "balance": None,
        "mode": "signed_amount",
    }

    result = loader.standardize_data(df, mapping, "bank", "auto", skiprows_offset=3)

    assert result["original_idx"].tolist() == [1]
    assert result["original_file_row"].tolist() == [5]


def test_参数校验覆盖全部数值输入():
    is_valid, message = validate_config_params(
        "31",
        "31",
        "30",
        "3",
        random_seed="abc",
        similarity_threshold="0.5",
        similarity_high="0.7",
        max_candidates="30",
        memory_limit="6.0",
        bank_skip="0",
        journal_skip="0",
    )

    assert is_valid is False
    assert "随机种子必须为整数" in message


def test_高相似度阈值不能小于基础阈值():
    is_valid, message = validate_config_params(
        "30",
        "30",
        "20",
        "3",
        random_seed="0",
        similarity_threshold="0.8",
        similarity_high="0.7",
        max_candidates="30",
        memory_limit="6.0",
        bank_skip="0",
        journal_skip="0",
    )

    assert is_valid is False
    assert "高相似度阈值不能小于相似度阈值" in message


def test_相似度阈值必须在零到一之间():
    is_valid, message = validate_config_params(
        "30",
        "30",
        "20",
        "3",
        random_seed="0",
        similarity_threshold="1.2",
        similarity_high="1.2",
        max_candidates="30",
        memory_limit="6.0",
        bank_skip="0",
        journal_skip="0",
    )

    assert is_valid is False
    assert "相似度阈值必须在 0 到 1 之间" in message


def test_相似度阈值允许零作为边界值():
    is_valid, message = validate_config_params(
        "30",
        "30",
        "20",
        "3",
        random_seed="0",
        similarity_threshold="0",
        similarity_high="0",
        max_candidates="30",
        memory_limit="6.0",
        bank_skip="0",
        journal_skip="0",
    )

    assert is_valid is True
    assert message == ""


def test_续行空日期会沿用上一行日期():
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    df = pd.DataFrame(
        {
            "日期": ["2024-01-01", "", "2024-01-02"],
            "金额": [100, 200, 300],
            "摘要": ["首行", "空日期", "末行"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "voucher": None,
        "balance": None,
        "mode": "signed_amount",
    }

    result = loader.standardize_data(df, mapping, "bank", "auto")

    assert result["date"].tolist() == [
        pd.Timestamp("2024-01-01"),
        pd.Timestamp("2024-01-01"),
        pd.Timestamp("2024-01-02"),
    ]
    assert collector.get_summary()["空日期行"] == 0


def test_首行空日期仍会记录为异常():
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    df = pd.DataFrame(
        {
            "日期": ["", "2024-01-02"],
            "金额": [100, 300],
            "摘要": ["无上文日期", "正常"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "voucher": None,
        "balance": None,
        "mode": "signed_amount",
    }

    result = loader.standardize_data(df, mapping, "bank", "auto")

    assert result["date"].tolist() == [pd.Timestamp("2024-01-02")]
    assert collector.get_summary()["空日期行"] == 1


def test_凭证号留空但借贷方有值时仍按续行补日期():
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    df = pd.DataFrame(
        {
            "日期": ["2024-01-01", "", "2024-01-02"],
            "凭证号": ["记-001", "", "记-002"],
            "摘要": ["首笔", "", "次日"],
            "借方": [Decimal("100.00"), Decimal("0.00"), Decimal("0.00")],
            "贷方": [Decimal("0.00"), Decimal("15.00"), Decimal("80.00")],
        }
    )
    mapping = {
        "date": "日期",
        "voucher": "凭证号",
        "summary": "摘要",
        "debit": "借方",
        "credit": "贷方",
        "balance": None,
        "mode": "debit_credit",
    }

    result = loader.standardize_data(df, mapping, "journal", "auto")

    # 用 original_idx 定位行（standardize_data 会按日期+金额排序，original_file_row 顺序不确定）
    result_by_row = result.set_index("original_idx")[["date", "amount"]]
    assert result_by_row.loc[1, "date"] == pd.Timestamp("2024-01-01")
    assert result_by_row.loc[2, "date"] == pd.Timestamp("2024-01-01")
    assert result_by_row.loc[3, "date"] == pd.Timestamp("2024-01-02")
    assert result_by_row.loc[1, "amount"] == Decimal("100.00")
    assert result_by_row.loc[2, "amount"] == Decimal("-15.00")
    assert result_by_row.loc[3, "amount"] == Decimal("-80.00")
    assert collector.get_summary()["空日期行"] == 0


def test_未知方向不会被静默当作收入():
    """真正无法识别的方向值必须被拒绝，而不是静默当正数"""
    collector = ParseErrorCollector()
    loader = DataLoader(error_collector=collector)
    df = pd.DataFrame(
        {
            "日期": ["2024-01-01", "2024-01-02"],
            "金额": ["100", "200"],
            "方向": ["收", "莫名其妙的方向"],
            "摘要": ["正常", "未知方向"],
        }
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "summary": "摘要",
        "voucher": None,
        "balance": None,
        "mode": "single_amount_with_direction",
    }

    result = loader.standardize_data(df, mapping, "bank", "auto")

    assert len(result) == 1
    assert result.iloc[0]["amount"] == Decimal("100.00")
    assert collector.get_summary()["方向解析失败"] == 1
    assert collector.get_summary()["金额解析失败"] == 0


def test_随机贪心未精确命中时必须返回空():
    result = _randomized_greedy(
        window_amounts=[5000, 2000],
        window_dates=[pd.Timestamp("2024-01-01"), pd.Timestamp("2024-01-01")],
        window_indices=[1, 2],
        target=6000,
        num_attempts=5,
        random_seed=0,
    )

    assert result is None


def test_随机贪心处理表格压缩整数时不会溢出():
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        result = _randomized_greedy(
            window_amounts=[
                np.uint32(50_000_000),
                np.uint32(10_000),
            ],
            window_dates=[
                pd.Timestamp("2024-01-01"),
                pd.Timestamp("2024-01-01"),
            ],
            window_indices=[1, 2],
            target=np.uint32(1_000),
            num_attempts=1,
            random_seed=0,
        )

    assert result is None
    assert not [
        warning
        for warning in caught
        if issubclass(warning.category, RuntimeWarning)
    ]


def test_日总额核销必须校验收支结构():
    bank = pd.DataFrame(
        {
            "date": [pd.Timestamp("2024-01-01"), pd.Timestamp("2024-01-01")],
            "summary": ["收款", "付款"],
            "amount_decimal": [10000, -4000],
        }
    )
    journal = pd.DataFrame(
        {
            "date": [pd.Timestamp("2024-01-01")],
            "summary": ["净额一笔"],
            "amount_decimal": [6000],
        }
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.match_daily_total()

    assert matcher.bank["matched"].sum() == 0
    assert matcher.journal["matched"].sum() == 0


def test_批量聚合匹配不会吞并普通重复流水():
    bank = pd.DataFrame(
        {
            "date": [pd.Timestamp("2024-01-01")] * 11,
            "summary": ["普通付款"] * 11,
            "amount_decimal": [100] * 11,
        }
    )
    journal = pd.DataFrame(
        {
            "date": [pd.Timestamp("2024-01-01")],
            "summary": ["普通汇总"],
            "amount_decimal": [1100],
        }
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.match_batch_aggregation()

    assert matcher.bank["matched"].sum() == 0
    assert matcher.journal["matched"].sum() == 0


def test_Windows控制台不支持日志图标时不会中断核对(monkeypatch):
    empty = pd.DataFrame(
        columns=["date", "summary", "amount_decimal"]
    )
    matcher = Matcher(empty, empty, MatcherConfig())
    printed = []

    def gbk_print(message):
        printed.append(message)
        if len(printed) == 1:
            raise UnicodeEncodeError(
                "gbk",
                str(message),
                0,
                1,
                "无法显示该字符",
            )

    monkeypatch.setattr("builtins.print", gbk_print)

    matcher._log("🧹 已完成")

    assert len(printed) == 2
    assert "已完成" in printed[-1]
