"""阶段一输入边界：硬冲突阻止匹配，金额原值与方向证据必须保留。"""
from decimal import Decimal

import pandas as pd
import pytest

from application import run_reconciliation
from data_loader import DataLoader
from data_structures import MatcherConfig
from input_precheck import InputPrecheckBlockedError, TableStructure, build_input_precheck


@pytest.mark.parametrize(
    ("银行账户", "账账户", "银行币种", "账币种", "检查项目"),
    [
        (["A001", "A001"], ["A002", "A002"], "CNY", "CNY", "核对账户"),
        (["A001", "A001"], ["A001", "A001"], "CNY", "USD", "核对币种"),
        (["A001", "A002"], ["A001", "A001"], "CNY", "CNY", "核对账户"),
    ],
    ids=["账户明确不同", "币种明确不同", "单文件混合账户"],
)
def test_真实入口对明确范围冲突停止匹配且不生成报告(
    tmp_path, 银行账户, 账账户, 银行币种, 账币种, 检查项目
):
    # 若范围冲突仅警告后继续，两笔同业务同金额会被自动确认并生成报告。
    for 文件名, 账户, 币种 in [("银行.csv", 银行账户, 银行币种), ("账.csv", 账账户, 账币种)]:
        pd.DataFrame({
            "日期": ["2026-08-10", "2026-08-11"],
            "金额": [100, 200],
            "摘要": ["甲公司回款", "乙公司回款"],
            "业务编号": ["P01", "P02"],
            "本方账号": 账户,
            "币种": [币种, 币种],
        }).to_csv(tmp_path / 文件名, index=False, encoding="utf-8-sig")
    映射 = {
        "date": "日期", "amount": "金额", "summary": "摘要", "mode": "signed_amount",
        "auxiliary_text_columns": ["摘要", "业务编号"],
    }
    报告 = tmp_path / "核对结果.xlsx"

    with pytest.raises(InputPrecheckBlockedError) as 异常:
        run_reconciliation(
            str(tmp_path / "银行.csv"), str(tmp_path / "账.csv"),
            映射, 映射, MatcherConfig(),
            bank_skiprows=0, journal_skiprows=0,
            bank_header_rows=1, journal_header_rows=1,
            output_path=报告,
        )

    assert 异常.value.report.has_blockers
    assert any(项.name == 检查项目 and 项.status == "无法计算" for 项 in 异常.value.report.items)
    assert not 报告.exists()


@pytest.mark.parametrize(
    ("原始金额字段", "金额映射", "期望证据"),
    [
        (
            {"借方": -100, "贷方": 0},
            {"mode": "debit_credit", "debit": "借方", "credit": "贷方"},
            {
                "mode": "debit_credit",
                "debit_column": "借方", "debit_value": -100,
                "credit_column": "贷方", "credit_value": 0,
                "amount_column": None, "amount_value": None,
                "direction_column": None, "direction_value": None,
            },
        ),
        (
            {"发生额": -100, "借贷方向": "借"},
            {"mode": "single_amount_with_direction", "amount": "发生额", "direction": "借贷方向"},
            {
                "mode": "single_amount_with_direction",
                "debit_column": None, "debit_value": None,
                "credit_column": None, "credit_value": None,
                "amount_column": "发生额", "amount_value": -100,
                "direction_column": "借贷方向", "direction_value": "借",
            },
        ),
    ],
    ids=["借贷分列红字来源", "单金额和方向来源"],
)
def test_标准化结果保留原始金额及方向证据(原始金额字段, 金额映射, 期望证据):
    # 仅保留净金额无法区分正常借方与贷方红字，后续冲销必须能读到原列及原值。
    原始 = pd.DataFrame([{"日期": "2026-08-10", "摘要": "红字冲销", **原始金额字段}])
    结果 = DataLoader().standardize_data(
        原始, {"date": "日期", "summary": "摘要", **金额映射}, "journal"
    )

    assert "amount_evidence" in 结果.columns
    证据 = 结果.iloc[0]["amount_evidence"]
    assert isinstance(证据, dict)
    assert set(期望证据) <= set(证据)
    assert {键: 证据[键] for 键 in 期望证据} == 期望证据


@pytest.mark.parametrize(
    ("来源", "方向", "期望金额"),
    [("journal", "借", Decimal("-100.00")), ("bank", "支出", Decimal("100.00"))],
    ids=["账侧借方负数", "银行支出负数"],
)
def test_单金额加方向按原有符号相乘而不取绝对值(来源, 方向, 期望金额):
    原始 = pd.DataFrame([{"日期": "2026-08-10", "发生额": -100, "方向": 方向, "摘要": "红字冲销"}])
    结果 = DataLoader().standardize_data(
        原始,
        {"date": "日期", "amount": "发生额", "direction": "方向", "summary": "摘要",
         "mode": "single_amount_with_direction"},
        来源,
    )

    assert 结果.iloc[0]["amount"] == 期望金额


@pytest.mark.parametrize(
    ("银行身份", "账身份", "期望账户状态", "期望币种状态"),
    [
        ({"本方账号": "12345", "币种": "CNY"},
         {"本方账号": 12345.0, "币种": "人民币"}, "通过", "通过"),
        ({"币种": "CNY"}, {"本方账号": "12345", "币种": "CNY"}, "疑点", "通过"),
        ({"本方账号": "12345", "币种": "CNY"}, {"币种": "CNY"}, "疑点", "通过"),
        ({"本方账号": "12345"}, {"本方账号": "12345", "币种": "CNY"}, "通过", "疑点"),
        ({"本方账号": "12345", "币种": "CNY"}, {"本方账号": "12345"}, "通过", "疑点"),
    ],
    ids=["账号数值尾零及人民币别名一致", "银行未提供账户", "账未提供账户",
         "银行未提供币种", "账未提供币种"],
)
def test_范围格式差异和身份信息缺失不得误阻断(
    银行身份, 账身份, 期望账户状态, 期望币种状态
):
    # 身份缺失只限制核对结论；不同展示形式也不能被误判为明确不同。
    原始银行 = pd.DataFrame([{"日期": "2026-08-10", "金额": 100,
                             "摘要": "甲公司回款", **银行身份}])
    原始账 = pd.DataFrame([{"日期": "2026-08-10", "金额": 100,
                           "摘要": "甲公司回款", **账身份}])
    映射 = {"date": "日期", "amount": "金额", "summary": "摘要", "mode": "signed_amount"}
    加载器 = DataLoader()
    检查 = build_input_precheck(
        raw_bank=原始银行, raw_journal=原始账,
        bank=加载器.standardize_data(原始银行, 映射, "bank"),
        journal=加载器.standardize_data(原始账, 映射, "journal"),
        bank_mapping=映射, journal_mapping=映射,
        bank_structure=TableStructure(0, 1, list(原始银行.columns), 10),
        journal_structure=TableStructure(0, 1, list(原始账.columns), 10),
    )

    按名称 = {项.name: 项 for 项 in 检查.items}
    assert 按名称["核对账户"].status == 期望账户状态
    assert 按名称["核对币种"].status == 期望币种状态
    assert not 检查.has_blockers
    for 名称, 期望状态 in [("核对账户", 期望账户状态), ("核对币种", 期望币种状态)]:
        if 期望状态 == "疑点":
            assert 按名称[名称].comparison == "范围身份未完全验证"
            assert "范围限制" in 按名称[名称].explanation
