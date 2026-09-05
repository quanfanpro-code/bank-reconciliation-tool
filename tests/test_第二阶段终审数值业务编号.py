"""第二阶段终审：数值型业务编号的读取、证据保留与匹配边界。"""

import pandas as pd

from data_loader import DataLoader
from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from reporter import Reporter
from 业务分组 import row_business


def _映射():
    return {
        "date": "日期",
        "summary": "摘要",
        "mode": "debit_credit",
        "debit": "借方金额",
        "credit": "贷方金额",
        "auxiliary_text_columns": ["摘要", "业务编号"],
    }


def _标准化(原表, 来源):
    return DataLoader().standardize_data(原表, _映射(), 来源)


def test_Excel数值型整数编号保留原值并与文本整数编号组成完整业务组():
    # Excel 引擎可能把数值编号交给 DataLoader 为 123.0；这里保留该原始类型，
    # 避免先手工转成字符串而掩盖真实问题。
    银行原表 = pd.DataFrame(
        [
            {
                "日期": "2026-08-10",
                "摘要": "项目结算汇总",
                "借方金额": 1000,
                "贷方金额": 0,
                "业务编号": 123.0,
            },
        ]
    )
    日记账原表 = pd.DataFrame([
        {
            "日期": "2026-08-10",
            "摘要": "项目结算明细甲",
            "借方金额": 0,
            "贷方金额": 400,
            "业务编号": "123",
        },
        {
            "日期": "2026-08-10",
            "摘要": "项目结算明细乙",
            "借方金额": 0,
            "贷方金额": 600,
            "业务编号": "123",
        },
    ])

    银行 = _标准化(银行原表, "bank")
    日记账 = _标准化(日记账原表, "journal")

    assert isinstance(银行原表.at[0, "业务编号"], float)
    assert 银行原表.at[0, "业务编号"] == 123.0
    assert 日记账原表.at[0, "业务编号"] == "123"
    assert 银行.iloc[0]["aux_text_fields"]["业务编号"] == "123.0"
    assert row_business(银行.iloc[0])["ids"] == {("业务", "123")}
    assert row_business(日记账.iloc[0])["ids"] == {("业务", "123")}

    实例 = Matcher(银行, 日记账, MatcherConfig(), logger=lambda _: None)
    实例.run()

    assert len(实例.selected_candidates) == 1
    候选 = 实例.selected_candidates[0]
    assert (候选.bank_idxs, 候选.journal_idxs) == ((0,), (0, 1))
    assert 候选.evidence.get("shared_business_id") is True
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
    组成 = Reporter(实例).build_report_tables(实例.config)["匹配组成"]
    银行组成 = 组成.loc[组成["来源"] == "银行流水"]
    assert len(银行组成) == 1
    assert "业务编号：123.0" in str(银行组成.iloc[0]["辅助文字"])


def test_Excel文本编号123点0不得按数值编号规则改成123():
    原表 = pd.DataFrame([{
        "日期": "2026-08-10",
        "摘要": "项目结算",
        "借方金额": 1000,
        "贷方金额": 0,
        "业务编号": "123.0",
    }])
    标准表 = _标准化(原表, "bank")

    assert 原表.at[0, "业务编号"] == "123.0"
    assert 标准表.iloc[0]["aux_text_fields"]["业务编号"] == "123.0"
    assert row_business(标准表.iloc[0])["ids"] == {("业务", "123.0")}
