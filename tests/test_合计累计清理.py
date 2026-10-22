"""汇总不能重复计入交易，正常业务和日期续行不能误删。"""
from decimal import Decimal

import pandas as pd
import pytest

from data_loader import DataLoader, ParseErrorCollector
from input_precheck import _non_transaction_mask


@pytest.mark.parametrize("来源", ["bank", "journal"])
@pytest.mark.parametrize("标签列", ["日期", "月份", "摘要", "其他说明"])
def test_合计累计位于不同列且凭证位置有笔数仍排除并保留原始行号(来源, 标签列):
    行 = [{"日期": "2026-09-01", "月份": 9, "借方": 100, "贷方": 0,
          "摘要": "正常业务", "凭证": 1, "其他说明": None}]
    for 标签, 笔数 in [(" 9月合计： ", None), ("本期累计:", None), ("合计：", 1)]:
        项 = {"日期": None, "月份": None, "借方": 100, "贷方": 0,
              "摘要": None, "凭证": 笔数, "其他说明": None}
        项[标签列] = 标签
        行.append(项)
    原表 = pd.DataFrame(行)
    映射 = {"date": "日期", "debit": "借方", "credit": "贷方", "summary": "摘要",
          "voucher": "凭证", "mode": "debit_credit"}
    异常 = ParseErrorCollector()
    结果 = DataLoader(error_collector=异常).standardize_data(
        原表, 映射, 来源, skiprows_offset=3,
    )

    assert len(结果) == 1
    assert 结果.original_file_row.tolist() == [5]
    assert abs(结果.amount.iloc[0]) == Decimal("100")
    排除 = [项 for 项 in 异常.get_all_errors() if 项["type"] == "被丢弃的汇总行"]
    assert [项["row"] for 项 in 排除] == [6, 7, 8]
    assert all(项["original_data"][标签列] == 行[序号][标签列]
               for 序号, 项 in enumerate(排除, 1))
    assert _non_transaction_mask(原表, 映射).tolist() == [False, True, True, True]
    assert len(原表) == 4


@pytest.mark.parametrize("来源", ["bank", "journal"])
def test_正常业务含合计累计和只有金额的日期续行仍保留(来源):
    原表 = pd.DataFrame([
        {"日期": "2026-09-01", "金额": 100, "摘要": "累计结算货款"},
        {"日期": None, "金额": 50, "摘要": "支付项目合计款"},
        {"日期": None, "金额": 25, "摘要": None},
    ])
    映射 = {"date": "日期", "amount": "金额", "summary": "摘要", "mode": "signed_amount"}
    结果 = DataLoader().standardize_data(原表, 映射, 来源)

    assert 结果.amount.sum() == Decimal("175")
    assert set(结果.original_file_row) == {2, 3, 4}
    assert 结果.date.nunique() == 1
    assert _non_transaction_mask(原表, 映射).tolist() == [False, False, False]


@pytest.mark.parametrize("来源", ["bank", "journal"])
def test_带有效日期的明确汇总与有凭证的正常交易分类一致(来源):
    原表 = pd.DataFrame([
        {"日期": "2026-09-01", "金额": 100, "摘要": "正常业务", "凭证": 1},
        {"日期": "2026-09-01", "金额": 100, "摘要": "本期合计", "凭证": None},
        {"日期": "2026-09-01", "金额": 50, "摘要": "本期合计", "凭证": 2},
    ])
    映射 = {"date": "日期", "amount": "金额", "summary": "摘要", "voucher": "凭证", "mode": "signed_amount"}
    结果 = DataLoader().standardize_data(原表, 映射, 来源)

    assert 结果.amount.sum() == Decimal("150")
    assert set(结果.original_file_row) == {2, 4}
    assert _non_transaction_mask(原表, 映射).tolist() == [False, True, False]


@pytest.mark.parametrize("来源", ["bank", "journal"])
def test_错误日期的业务文字不能掩盖为合计累计行(来源):
    原表 = pd.DataFrame([
        {"日期": "2026-09-01", "金额": 100, "摘要": "正常业务"},
        {"日期": "错误日期", "金额": 50, "摘要": "累计结算货款"},
    ])
    映射 = {"date": "日期", "amount": "金额", "summary": "摘要", "mode": "signed_amount"}
    异常 = ParseErrorCollector()
    结果 = DataLoader(error_collector=异常).standardize_data(原表, 映射, 来源)

    assert len(结果) == 1
    assert 异常.get_summary()["日期解析失败"] == 1
    assert 异常.get_summary()["被丢弃的汇总行"] == 0
    assert _non_transaction_mask(原表, 映射).tolist() == [False, False]


@pytest.mark.parametrize("来源", ["bank", "journal"])
def test_没有摘要映射的其他列累计标签也不进入交易(来源):
    原表 = pd.DataFrame([
        {"日期": "2026-09-01", "金额": 100, "月份": 9},
        {"日期": None, "金额": 100, "月份": "本期累计："},
    ])
    映射 = {"date": "日期", "amount": "金额", "mode": "signed_amount"}
    结果 = DataLoader().standardize_data(原表, 映射, 来源)

    assert 结果.amount.sum() == Decimal("100")
    assert 结果.original_file_row.tolist() == [2]
