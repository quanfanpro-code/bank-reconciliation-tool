"""余额坏值即使不排除交易行，也必须降低总体范围和具体关系。"""
import pandas as pd

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig, ProcessingStatus, RiskLevel
from input_precheck import TableStructure, build_input_precheck
from matcher import Matcher
from reporter import Reporter


def _结构(frame):
    return TableStructure(0, 1, list(frame.columns), len(frame))


def test_有效交易行余额解析失败不得被人口分类掩盖():
    银行原表 = pd.DataFrame({
        "日期": ["2026-08-10", "2026-08-11", "2026-08-12"],
        "金额": [100, 200, 300],
        "余额": ["余额坏值", 1300, 1600],
        "摘要": ["项目甲收款", "项目乙收款", "项目丙收款"],
        "账号": ["A001"] * 3,
        "币种": ["CNY"] * 3,
    })
    日记账原表 = 银行原表.copy()
    日记账原表["余额"] = [None, 1300, 1600]
    映射 = {
        "date": "日期",
        "amount": "金额",
        "balance": "余额",
        "summary": "摘要",
        "account": "账号",
        "currency": "币种",
        "amount_basis": "本位币",
        "mode": "signed_amount",
    }
    收集器 = ParseErrorCollector()
    加载器 = DataLoader(error_collector=收集器)
    银行 = 加载器.standardize_data(银行原表, 映射, "bank")
    日记账 = 加载器.standardize_data(日记账原表, 映射, "journal")
    预检 = build_input_precheck(
        raw_bank=银行原表,
        raw_journal=日记账原表,
        bank=银行,
        journal=日记账,
        bank_mapping=映射,
        journal_mapping=映射,
        bank_structure=_结构(银行原表),
        journal_structure=_结构(日记账原表),
        parse_errors=收集器.get_all_errors(),
    )

    assert sum(项目["type"] == "余额解析失败" for 项目 in 收集器.get_all_errors()) == 1
    assert 预检.overall_control.bank_balance_status == "通过"
    assert 预检.overall_control.journal_balance_status == "通过"
    assert 预检.overall_control.scope_limited is True
    assert any("余额解析失败" in 原因 for 原因 in 预检.overall_control.reasons)
    总体余额项 = next(项目 for 项目 in 预检.items if 项目.name == "总体余额控制")
    assert 总体余额项.status == "疑点"
    assert "余额解析失败" in 总体余额项.explanation

    匹配器 = Matcher(
        银行,
        日记账,
        MatcherConfig(),
        overall_control=预检.overall_control,
        logger=lambda _: None,
    )
    匹配器.run()
    assert 匹配器.selected_candidates
    assert all(
        候选.processing_status is ProcessingStatus.FLAGGED
        and 候选.risk_level is RiskLevel.UNKNOWN
        for 候选 in 匹配器.selected_candidates
    )
    系统结论 = Reporter(
        匹配器,
        precheck_report=预检,
    ).build_report_tables(匹配器.config)["核对结论"]
    结论文字 = str(
        系统结论.loc[系统结论["项目"] == "系统结论", "数值"].iloc[0]
    )
    assert "总体余额控制异常" in 结论文字
