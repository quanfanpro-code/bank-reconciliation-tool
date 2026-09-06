import pandas as pd

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig, ProcessingStatus, RiskLevel
from input_precheck import TableStructure, build_input_precheck
from matcher import Matcher


映射 = {
    "date": "日期",
    "amount": "金额",
    "summary": "摘要",
    "mode": "signed_amount",
    "amount_basis": "本位币",
    "account": "账号",
    "currency": "币种",
    "auxiliary_text_columns": ["摘要"],
}


def _结构(原表):
    return TableStructure(0, 1, list(原表.columns), 10)


def test_被日期解析异常排除的有金额原行必须降低具体关系结论():
    原银行 = pd.DataFrame(
        {
            "日期": ["坏日期", "2026-08-31"],
            "金额": [50, 100],
            "摘要": ["同一项目回款", "同一项目回款"],
            "账号": ["A001", "A001"],
            "币种": ["CNY", "CNY"],
        }
    )
    原日记账 = pd.DataFrame(
        {
            "日期": ["2026-08-31"],
            "金额": [100],
            "摘要": ["同一项目回款"],
            "账号": ["A001"],
            "币种": ["CNY"],
        }
    )
    收集器 = ParseErrorCollector()
    加载器 = DataLoader(error_collector=收集器)
    银行 = 加载器.standardize_data(原银行, 映射, "bank")
    日记账 = 加载器.standardize_data(原日记账, 映射, "journal")

    预检 = build_input_precheck(
        raw_bank=原银行,
        raw_journal=原日记账,
        bank=银行,
        journal=日记账,
        bank_mapping=映射,
        journal_mapping=映射,
        bank_structure=_结构(原银行),
        journal_structure=_结构(原日记账),
        parse_errors=收集器.get_all_errors(),
    )

    assert 预检.overall_control.scope_limited is True
    assert any("数据入口" in 原因 and "解析异常" in 原因 for 原因 in 预检.overall_control.reasons)

    匹配器 = Matcher(
        银行,
        日记账,
        MatcherConfig(),
        overall_control=预检.overall_control,
        logger=lambda _: None,
    )
    匹配器.run()

    assert len(匹配器.selected_candidates) == 1
    候选 = 匹配器.selected_candidates[0]
    assert 候选.processing_status is ProcessingStatus.FLAGGED
    assert 候选.risk_level is RiskLevel.UNKNOWN
    assert "总体资料尚未闭合" in 候选.processing_reason
    assert "数据入口存在解析异常" in 候选.processing_reason
