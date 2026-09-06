"""输入完整性独立复核：余额覆盖、显式范围、逐行人口与阻断证据。"""
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

from application import run_reconciliation
from balance import build_overall_controls, check_balance_continuity
from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig
from input_precheck import InputPrecheckBlockedError, TableStructure, build_input_precheck


映射 = {"date": "日期", "amount": "金额", "mode": "signed_amount", "summary": "摘要",
        "voucher": "凭证号", "amount_basis": "本位币"}


def 原表():
    return pd.DataFrame({"日期": ["2026-08-10", "2026-08-11"], "金额": [100, 200],
                         "摘要": ["项目回款", "项目回款"], "凭证号": ["记001", "记002"],
                         "本方账号": ["A001", "A001"], "币种": ["CNY", "CNY"]})


def 预检(银行, 账, 银行映射=None, 标准银行=None, 错误=()):
    加载器 = DataLoader()
    return build_input_precheck(raw_bank=银行, raw_journal=账,
        bank=标准银行 if 标准银行 is not None else 加载器.standardize_data(银行, 银行映射 or 映射, "bank"),
        journal=加载器.standardize_data(账, 映射, "journal"),
        bank_mapping=银行映射 or 映射, journal_mapping=映射,
        bank_structure=TableStructure(0, 1, list(银行.columns), 10),
        journal_structure=TableStructure(0, 1, list(账.columns), 10), parse_errors=错误)


def test_单个余额点只能说明覆盖不足不能声称连续性通过():
    标准 = pd.DataFrame({"date": [pd.Timestamp("2026-08-10")], "amount": [Decimal("100")],
                         "balance": [Decimal("1100")], "original_file_row": [2]})
    控制 = build_overall_controls(标准, 标准)
    assert 控制.bank_balance_status != "通过"
    assert any("余额点" in 原因 or "不足" in 原因 or "未实施" in 原因 for 原因 in 控制.reasons)


def test_逐原行余额连续性能够识别同日中间错误并保留原行():
    标准 = pd.DataFrame({"date": pd.to_datetime(["2026-08-10"] * 3),
                         "amount": [Decimal("100")] * 3,
                         "balance": [Decimal("1100"), Decimal("1250"), Decimal("1300")],
                         "original_file_row": [2, 3, 4]})
    异常 = check_balance_continuity(标准.iloc[[2, 0, 1]], source="银行流水")
    assert len(异常) == 2
    assert [项["差额"] for 项 in 异常] == [Decimal("50"), Decimal("50")]
    assert [项["原文件行号"] for 项 in 异常] == [3, 4]


def test_同日正确余额与跨空余额区间均不误报异常():
    标准 = pd.DataFrame({"date": pd.to_datetime(["2026-08-10"] * 3),
                         "amount": [Decimal("100")] * 3,
                         "balance": [Decimal("1100"), None, Decimal("1300")],
                         "original_file_row": [2, 3, 4]})
    assert check_balance_continuity(标准) == []
    assert build_overall_controls(标准, 标准).bank_balance_status == "通过"


@pytest.mark.parametrize(("键", "名称"), [("account", "核对账户"), ("currency", "核对币种")])
def test_显式范围列不存在必须阻断而不能扫描另一个同名业务列(键, 名称):
    检查 = 预检(原表(), 原表(), {**映射, 键: "已被删除的列"})
    项 = next(项 for 项 in 检查.items if 项.name == 名称)
    assert 项.status == "无法计算"
    assert "已被删除的列" in 项.explanation


def test_人口摘要以唯一分类计数且各类金额与原始可解析金额闭合():
    银行 = 原表()
    银行.loc[2] = ["无效日期", 300, "待查回款", "记003", "A001", "CNY"]
    银行.loc[3] = ["2026-08-11", 600, "合计", "", "A001", "CNY"]
    收集器 = ParseErrorCollector()
    标准 = DataLoader(error_collector=收集器).standardize_data(银行, 映射, "bank")
    检查 = 预检(银行, 原表(), 标准银行=标准, 错误=收集器.get_all_errors())
    明细 = 检查.population_dataframe().query("来源 == '银行流水'")
    assert 明细["原文件行号"].is_unique
    assert 明细["唯一处置类别"].tolist() == ["有效交易", "有效交易", "解析异常", "非交易行"]
    assert 明细["原始可解析净额"].tolist() == [Decimal("100"), Decimal("200"), Decimal("300"), Decimal("600")]
    项 = next(项 for 项 in 检查.items if 项.name == "数据入口")
    for 文字 in ["原始4行", "有效交易2行", "非交易1行", "解析异常1行", "其他排除0行",
                 "原始可解析净额1200.00", "有效交易净额300.00", "非交易行净额600.00",
                 "解析异常净额300.00", "其他排除净额0.00", "金额勾稽差额0.00"]:
        assert 文字 in 项.bank_result


def test_已进入标准化的有效行不能再计为非交易导致人口不闭合():
    银行 = 原表()
    标准 = DataLoader().standardize_data(银行, 映射, "bank")
    # 有效行优先级已确定，原始行的文字规则不能再使摘要计数重叠。
    银行.loc[0, "摘要"] = "合计"
    检查 = 预检(银行, 原表(), 标准银行=标准)
    项 = next(项 for 项 in 检查.items if 项.name == "数据入口")
    assert "原始2行；有效交易2行；非交易0行" in 项.bank_result


def test_无法解析金额必须保留原值说明而不能列作零金额():
    银行 = 原表().astype({"金额": object})
    银行.loc[1, "金额"] = "金额不明"
    收集器 = ParseErrorCollector()
    标准 = DataLoader(error_collector=收集器).standardize_data(银行, 映射, "bank")
    检查 = 预检(银行, 原表(), 标准银行=标准, 错误=收集器.get_all_errors())
    异常 = 检查.population_dataframe().query("来源 == '银行流水' and 唯一处置类别 == '解析异常'").iloc[0]
    assert pd.isna(异常["原始可解析净额"])
    assert "金额不明" in 异常["原金额值"]
    assert "无法解析" in 异常["金额去向说明"]


def test_标准化方向阻断的问题报告保存逐行解析错误(tmp_path):
    银行, 账 = 原表(), 原表()
    银行["方向"] = ["未知方向", "无法识别方向"]
    for 名称, 数据 in [("银行.xlsx", 银行), ("序时账.xlsx", 账)]:
        数据.to_excel(tmp_path / 名称, index=False)
    with pytest.raises(InputPrecheckBlockedError) as 错:
        run_reconciliation(str(tmp_path / "银行.xlsx"), str(tmp_path / "序时账.xlsx"),
            {**映射, "mode": "single_amount_with_direction", "direction": "方向"}, 映射,
            MatcherConfig(), bank_skiprows=0, journal_skiprows=0, bank_header_rows=1,
            journal_header_rows=1, output_path=tmp_path / "正式报告.xlsx")
    assert not (tmp_path / "正式报告.xlsx").exists()
    assert 错.value.report_path
    表 = pd.read_excel(Path(错.value.report_path), sheet_name=None)
    assert "解析异常明细" in 表
    文字 = " ".join(表["解析异常明细"].fillna("").astype(str).to_numpy().ravel())
    assert "无法识别方向" in 文字 and "方向解析失败" in 文字
    assert "3" in 文字
