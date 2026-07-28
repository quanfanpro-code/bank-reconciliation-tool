from decimal import Decimal

import pandas as pd

from application import run_reconciliation
from data_structures import MatcherConfig
from llm_assistant import LLMConfig


def _build_business_workbooks(tmp_path):
    bank_rows = [
        {
            "日期": "2026-01-15",
            "金额": value,
            "方向": "贷",
            "摘要": "批次核销",
        }
        for value in (200, 200, 100)
    ]
    journal_rows = [
        {
            "日期": "2026-01-15",
            "金额": 100,
            "方向": "借",
            "摘要": "批次核销",
        }
        for _ in range(5)
    ]
    for index in range(1, 22):
        bank_rows.append(
            {
                "日期": f"2026-02-{index:02d}",
                "金额": 5000,
                "方向": "贷",
                "摘要": f"微小差异项目{index}",
            }
        )
        journal_rows.append(
            {
                "日期": f"2026-02-{index:02d}",
                "金额": 1,
                "方向": "借",
                "摘要": f"微小差异项目{index}",
            }
        )
    bank_rows.extend(
        [
            {
                "日期": "2026-03-22",
                "金额": 120000,
                "方向": "贷",
                "摘要": "重大设备采购",
            },
            {
                "日期": "2026-03-31",
                "金额": 999,
                "方向": "贷",
                "摘要": "银行独有记录",
            },
        ]
    )
    journal_rows.append(
        {
            "日期": "2026-03-22",
            "金额": 120000,
            "方向": "借",
            "摘要": "重大设备采购",
        }
    )

    bank_path = tmp_path / "银行.xlsx"
    journal_path = tmp_path / "日记账.xlsx"
    pd.DataFrame(bank_rows).to_excel(bank_path, index=False)
    pd.DataFrame(journal_rows).to_excel(journal_path, index=False)
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "summary": "摘要",
        "auxiliary_text_columns": ["摘要"],
        "mode": "single_amount_with_direction",
    }
    return bank_path, journal_path, mapping


def _numeric_metric(summary, name):
    values = summary.loc[summary["项目"] == name, "数值"]
    assert len(values) == 1
    return float(values.iloc[0])


def test_完整核对流程同时产生自动确认整池复核和未找到候选(
    tmp_path,
):
    bank_path, journal_path, mapping = _build_business_workbooks(
        tmp_path
    )
    output_path = run_reconciliation(
        bank_path=str(bank_path),
        journal_path=str(journal_path),
        bank_mapping=mapping,
        journal_mapping=mapping,
        matcher_config=MatcherConfig(
            performance_materiality=Decimal("100000"),
            clearly_trivial_threshold=Decimal("5000"),
            auto_confirm_score=70,
        ),
    )

    summary = pd.read_excel(output_path, sheet_name="核对汇总")
    pending = pd.read_excel(output_path, sheet_name="待人工复核")
    groups = pd.read_excel(output_path, sheet_name="匹配明细")
    unmatched = pd.read_excel(output_path, sheet_name="银行未达")

    assert 0 <= _numeric_metric(summary, "精确匹配率") <= 1
    assert 0 <= _numeric_metric(summary, "自动处理率") <= 1
    assert "自动确认" in set(groups["最终状态"])
    assert "月度累计超出实际执行重要性水平" in set(
        pending["原因"]
    )
    assert "重大设备采购" in set(
        pending.loc[pending["事项类型"] == "匹配组", "原因"].map(
            lambda value: "重大设备采购"
            if "超过实际执行重要性水平" in str(value)
            else ""
        )
    )
    assert "银行独有记录" in set(unmatched["摘要"])


def test_大模型超时自动降级但仍生成完整报告(tmp_path, monkeypatch):
    bank_path = tmp_path / "银行.xlsx"
    journal_path = tmp_path / "日记账.xlsx"
    pd.DataFrame(
        [
            {
                "日期": "2026-04-01",
                "金额": 1000,
                "方向": "贷",
                "摘要": "甲项目",
            },
            {
                "日期": "2026-04-01",
                "金额": 1000,
                "方向": "贷",
                "摘要": "乙项目",
            },
        ]
    ).to_excel(bank_path, index=False)
    pd.DataFrame(
        [
            {
                "日期": "2026-04-01",
                "金额": 1000,
                "方向": "借",
                "摘要": "甲项目",
            },
            {
                "日期": "2026-04-01",
                "金额": 1000,
                "方向": "借",
                "摘要": "乙项目",
            },
        ]
    ).to_excel(journal_path, index=False)
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "summary": "摘要",
        "auxiliary_text_columns": ["摘要"],
        "mode": "single_amount_with_direction",
    }

    class _TimeoutAssistant:
        def __init__(self, config):
            self.config = config
            self.candidate_limit = config.candidate_limit

        def evaluate_candidates(self, _request):
            raise TimeoutError("模拟大模型超时")

    monkeypatch.setattr(
        "application.LLMAssistant",
        _TimeoutAssistant,
    )
    output_path = run_reconciliation(
        bank_path=str(bank_path),
        journal_path=str(journal_path),
        bank_mapping=mapping,
        journal_mapping=mapping,
        matcher_config=MatcherConfig(),
        llm_config=LLMConfig(
            enabled=True,
            mode="online",
            base_url="https://example.com/v1",
            model="test-model",
        ),
    )

    workbook = pd.ExcelFile(output_path)
    assert "大模型辅助明细" in workbook.sheet_names
    llm_details = pd.read_excel(
        output_path,
        sheet_name="大模型辅助明细",
    )
    assert (llm_details["是否降级"] == "是").all()
    assert llm_details["错误"].str.contains("模拟大模型超时").any()
