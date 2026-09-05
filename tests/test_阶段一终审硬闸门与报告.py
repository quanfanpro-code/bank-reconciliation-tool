"""第一阶段终审：自动映射、账户硬闸门和总体控制报告必须同源。"""
from decimal import Decimal
from pathlib import Path

import pandas as pd
import pytest

from application import run_reconciliation
from balance import build_overall_controls
from data_loader import DataLoader
from data_structures import MatcherConfig, OverallControlResult
from gui import _auto_mapping_for_columns, auto_select_auxiliary_columns
from input_precheck import InputPrecheckBlockedError, TableStructure, build_input_precheck
from matcher import Matcher
from reporter import Reporter


def _结构(列名, 行数=1):
    return TableStructure(0, 1, list(列名), 行数)


def test_账户名称不得遮住真实账号且跨账户必须在匹配前阻断(tmp_path):
    银行 = pd.DataFrame({
        "日期": ["2026-08-10"],
        "账户名称": ["测试户"],
        "账号": ["A001"],
        "金额": [100],
        "摘要": ["项目回款"],
    })
    序时账 = 银行.copy()
    序时账["账号"] = "A002"
    银行映射 = _auto_mapping_for_columns(银行.columns, is_bank=True)
    账映射 = _auto_mapping_for_columns(序时账.columns, is_bank=False)

    assert 银行映射["account"] == 账映射["account"] == "账号"
    加载器 = DataLoader()
    检查 = build_input_precheck(
        raw_bank=银行,
        raw_journal=序时账,
        bank=加载器.standardize_data(银行, 银行映射, "bank"),
        journal=加载器.standardize_data(序时账, 账映射, "journal"),
        bank_mapping=银行映射,
        journal_mapping=账映射,
        bank_structure=_结构(银行.columns),
        journal_structure=_结构(序时账.columns),
    )
    账户检查 = next(项目 for 项目 in 检查.items if 项目.name == "核对账户")
    assert 检查.has_blockers is True
    assert 账户检查.status == "无法计算"

    银行路径 = tmp_path / "银行.xlsx"
    账路径 = tmp_path / "序时账.xlsx"
    银行.to_excel(银行路径, index=False)
    序时账.to_excel(账路径, index=False)
    正式报告 = tmp_path / "正式报告.xlsx"
    with pytest.raises(InputPrecheckBlockedError) as 异常:
        run_reconciliation(
            str(银行路径),
            str(账路径),
            银行映射,
            账映射,
            MatcherConfig(),
            bank_skiprows=0,
            journal_skiprows=0,
            bank_header_rows=1,
            journal_header_rows=1,
            output_path=正式报告,
        )
    assert not 正式报告.exists()
    assert 异常.value.problem_report_path
    assert Path(异常.value.problem_report_path).exists()


def test_主体同义词由界面自动保留且明确不同不得自动确认():
    主体字段 = ["对方名称", "客户名称", "供应商户名"]
    assert auto_select_auxiliary_columns(主体字段) == 主体字段

    银行原表 = pd.DataFrame({
        "日期": ["2026-08-10"],
        "金额": [100],
        "摘要": ["项目回款"],
        "对方名称": ["甲公司"],
    })
    账原表 = 银行原表.copy()
    账原表["对方名称"] = "乙公司"
    银行映射 = _auto_mapping_for_columns(银行原表.columns, is_bank=True)
    账映射 = _auto_mapping_for_columns(账原表.columns, is_bank=False)
    加载器 = DataLoader()
    匹配器 = Matcher(
        加载器.standardize_data(银行原表, 银行映射, "bank"),
        加载器.standardize_data(账原表, 账映射, "journal"),
        MatcherConfig(),
        logger=lambda _: None,
    )
    匹配器.run()

    assert all(
        候选.processing_status.value not in {"自动确认", "整组勾稽一致"}
        for 候选 in 匹配器.selected_candidates
    )
    assert any(
        "对方户名" in 候选.evidence.get("business_conflicts", ())
        for 候选 in 匹配器.candidates
    )


@pytest.mark.parametrize("模式", ["debit_credit", "single_amount_with_direction", "signed_amount"])
def test_报告期初余额必须直接复用前置总体控制且覆盖三种金额模式(模式):
    公共 = {"日期": ["2026-08-10"], "摘要": ["项目回款"], "余额": [1100]}
    if 模式 == "debit_credit":
        银行原表 = pd.DataFrame({**公共, "借方": [0], "贷方": [100]})
        账原表 = pd.DataFrame({**公共, "借方": [100], "贷方": [0]})
        映射 = {"date": "日期", "summary": "摘要", "balance": "余额",
              "debit": "借方", "credit": "贷方", "mode": 模式}
        银行映射 = 账映射 = 映射
    elif 模式 == "single_amount_with_direction":
        银行原表 = pd.DataFrame({**公共, "金额": [100], "方向": ["贷"]})
        账原表 = pd.DataFrame({**公共, "金额": [100], "方向": ["借"]})
        银行映射 = {"date": "日期", "summary": "摘要", "balance": "余额",
                    "amount": "金额", "direction": "方向", "mode": 模式}
        账映射 = dict(银行映射)
    else:
        银行原表 = pd.DataFrame({**公共, "金额": [100]})
        账原表 = 银行原表.copy()
        银行映射 = {"date": "日期", "summary": "摘要", "balance": "余额",
                    "amount": "金额", "mode": 模式}
        账映射 = dict(银行映射)

    加载器 = DataLoader()
    银行 = 加载器.standardize_data(银行原表, 银行映射, "bank")
    账 = 加载器.standardize_data(账原表, 账映射, "journal")
    总体控制 = build_overall_controls(银行, 账)
    assert 总体控制.bank_initial_balance == Decimal("1000")
    assert 总体控制.journal_initial_balance == Decimal("1000")
    匹配器 = Matcher(银行, 账, MatcherConfig(), overall_control=总体控制)
    匹配器.run()
    报告器 = Reporter(
        匹配器,
        raw_bank=银行原表,
        raw_journal=账原表,
        bank_mapping=银行映射,
        journal_mapping=账映射,
    )
    表 = 报告器.build_report_tables(匹配器.config)["核对结论"]
    结论 = dict(zip(表["项目"], 表["数值"]))

    assert 结论["银行期初余额"] == 1000
    assert 结论["日记账期初余额"] == 1000
    assert 结论["期初余额差额"] == 0
    assert 结论["期初余额状态"] == "一致"


def test_仅期间或人口受限时不得误写为总体余额控制异常():
    表 = pd.DataFrame({
        "date": [pd.Timestamp("2026-08-10")],
        "amount": [Decimal("100")],
        "amount_decimal": [1000000],
        "summary": ["项目回款"],
        "aux_text_fields": [{"摘要": "项目回款"}],
        "original_idx": [1],
        "original_file_row": [2],
    })
    总体控制 = OverallControlResult(
        period_status="疑点",
        amount_status="通过",
        scope_limited=True,
        reasons=("双方起止日期不一致",),
    )
    匹配器 = Matcher(表, 表.copy(), MatcherConfig(), overall_control=总体控制)
    匹配器.run()
    assert 匹配器.balance_integrity_limited is False
    assert 匹配器.overall_scope_limited is True
    核对结论 = Reporter(匹配器).build_report_tables(匹配器.config)["核对结论"]
    系统结论 = str(核对结论.loc[核对结论["项目"] == "系统结论", "数值"].iloc[0])
    assert "总体资料" in 系统结论 and "尚未闭合" in 系统结论
    assert "总体余额控制异常" not in 系统结论
