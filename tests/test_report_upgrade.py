from decimal import Decimal

import pandas as pd
from openpyxl import load_workbook

from data_structures import LLMDecisionRecord, MatcherConfig
from input_precheck import InputPrecheckReport, PrecheckItem
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter


def _report_df(date, amounts, source, summaries=None):
    summaries = summaries or [""] * len(amounts)
    records = [
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": summaries[index],
                "aux_text_fields": {"摘要": summaries[index]},
                "balance": None,
                "voucher_no": f"记-{index + 1:03d}" if source == "journal" else "",
                "source": source,
                "original_idx": index + 1,
                "original_file_row": index + 2,
            }
            for index, amount in enumerate(amounts)
        ]
    return pd.DataFrame(
        records,
        columns=[
            "date",
            "amount",
            "amount_decimal",
            "summary",
            "aux_text_fields",
            "balance",
            "voucher_no",
            "source",
            "original_idx",
            "original_file_row",
        ],
    )


def _three_vs_five_reporter():
    bank = _report_df(
        "2026-01-15",
        [200, 200, 100],
        "bank",
        ["销售回款"] * 3,
    )
    journal = _report_df(
        "2026-01-15",
        [100, 100, 100, 100, 100],
        "journal",
        ["销售回款"] * 5,
    )
    for frame in (bank, journal):
        frame["aux_text_fields"] = [
            {**fields, "批次号": "SALE-202601-01"}
            for fields in frame["aux_text_fields"]
        ]
    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
    )
    matcher.run()
    return Reporter(matcher)


def _pending_reporter():
    matcher = Matcher(
        _report_df(
            "2026-01-15",
            [120000],
            "bank",
            ["设备款"],
        ),
        _report_df(
            "2026-01-15",
            [120000, 120000],
            "journal",
            ["设备款", "设备款"],
        ),
        MatcherConfig(),
    )
    matcher.run()
    return Reporter(matcher)


def test_多对多报告按组展示且组成明细不伪造一一对应(tmp_path):
    output = tmp_path / "report.xlsx"
    _three_vs_five_reporter().generate_report(
        str(output),
        config=MatcherConfig(),
    )

    groups = pd.read_excel(output, sheet_name="整组勾稽")
    components = pd.read_excel(output, sheet_name="匹配组成")

    assert len(groups) == 1
    assert len(components) == 8
    assert set(components["来源"]) == {"银行流水", "日记账"}
    assert pd.api.types.is_numeric_dtype(components["金额"])


def test_匹配组成以正数展示金额并单列收支方向():
    matcher = Matcher(
        _report_df("2026-01-15", [-30], "bank", ["支付货款"]),
        _report_df("2026-01-15", [-30], "journal", ["支付货款"]),
        MatcherConfig(),
    )
    matcher.run()

    components = Reporter(matcher).build_report_tables(
        MatcherConfig()
    )["匹配组成"]

    assert "收支方向" in components.columns
    assert set(components["金额"]) == {30.0}
    assert set(components["收支方向"]) == {"支出"}


def test_双方余额列全空时报告不生成余额差异表():
    reporter = Reporter(
        Matcher(
            _report_df("2026-01-15", [70], "bank", ["单边记录"]),
            _report_df("2026-01-15", [], "journal"),
            MatcherConfig(),
        )
    )

    tables = reporter.build_report_tables(MatcherConfig())
    summary = dict(zip(tables["核对结论"]["项目"], tables["核对结论"]["数值"]))

    assert "余额差异明细" not in tables
    assert summary["余额核对"] == "余额核对未实施"


def test_报告不依赖人工填写也有完整系统结论():
    tables = _pending_reporter().build_report_tables(MatcherConfig())
    issues = tables["疑点事项"]

    assert {"系统结论", "风险等级", "判断依据", "建议动作"}.issubset(
        issues.columns
    )
    assert issues["系统结论"].notna().all()
    assert "待人工复核" not in set(issues["系统结论"])


def test_日总额整组不计入逐笔精确匹配率():
    tables = _three_vs_five_reporter().build_report_tables(MatcherConfig())
    summary = dict(zip(tables["核对结论"]["项目"], tables["核对结论"]["数值"]))

    assert summary["逐笔精确匹配率"] == 0
    assert summary["组级勾稽率"] == 1


def test_技术列和技术工作表保留但默认隐藏(tmp_path):
    output = tmp_path / "report.xlsx"
    _three_vs_five_reporter().generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output)
    sheet = workbook["整组勾稽"]
    headers = {cell.value: cell.column_letter for cell in sheet[1]}

    assert sheet.column_dimensions[headers["候选ID"]].hidden is True
    assert sheet.column_dimensions[headers["匹配ID"]].hidden is False
    assert workbook["运行参数"].sheet_state == "hidden"


def test_首页结论和输入说明不会因列宽过窄被截断(tmp_path):
    output = tmp_path / "report.xlsx"
    reporter = _pending_reporter()
    reporter.precheck_report = InputPrecheckReport(
        items=(
            PrecheckItem(
                name="余额可用性",
                bank_result="未提供",
                journal_result="未提供",
                comparison="不执行余额勾稽",
                status="提示",
                explanation=(
                    "余额列为空不影响逐笔与组合核对，报告自动略过余额差异表"
                ),
            ),
        )
    )
    reporter.generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output)

    summary = workbook["核对结论"]
    summary_headers = {
        cell.value: cell.column_letter for cell in summary[1]
    }
    assert summary.column_dimensions[summary_headers["数值"]].width >= 52
    assert summary.row_dimensions[2].height >= 30

    precheck = workbook["输入检查"]
    precheck_headers = {
        cell.value: cell.column_letter for cell in precheck[1]
    }
    assert precheck.column_dimensions[precheck_headers["说明"]].width >= 42
    assert precheck.row_dimensions[2].height >= 30


def test_人工字段只是允许空白的可选后续标注(tmp_path):
    output = tmp_path / "report.xlsx"
    _pending_reporter().generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output)
    sheet = workbook["疑点事项"]
    headers = [cell.value for cell in sheet[1]]
    validations = list(sheet.data_validations.dataValidation)

    assert {"后续状态", "处理说明", "调整凭证号", "责任人", "处理日期"}.issubset(headers)
    assert any(item.allow_blank for item in validations)
    status_column = headers.index("后续状态") + 1
    assert sheet.cell(row=2, column=status_column).value in (None, "")


def test_疑点事项提供可空白的后续状态和说明列(tmp_path):
    output = tmp_path / "report.xlsx"
    _pending_reporter().generate_report(
        str(output),
        config=MatcherConfig(),
    )

    workbook = load_workbook(output)
    sheet = workbook["疑点事项"]
    validations = list(sheet.data_validations.dataValidation)
    headers = [cell.value for cell in sheet[1]]

    assert '"已关注,已处理,无需处理"' in {
        item.formula1 for item in validations
    }
    assert "处理说明" in headers
    conclusion_column = headers.index("后续状态") + 1
    assert sheet.cell(row=2, column=conclusion_column).value in (None, "")
    assert sheet.auto_filter.ref is not None
    assert sheet.auto_filter.ref.startswith("B1:")


def test_疑点事项直接提供系统判断和核心证据():
    pending = _pending_reporter().build_report_tables(
        MatcherConfig()
    )["疑点事项"]

    assert {
        "银行原文件行号",
        "日记账原文件行号",
        "银行日期",
        "日记账日期",
        "银行金额",
        "日记账金额",
        "系统结论",
        "风险等级",
        "判断依据",
        "建议动作",
        "金额分",
        "日期分",
        "文字分",
        "结构分",
        "差异池ID",
    }.issubset(pending.columns)
    row = pending.iloc[0]
    assert "2" in str(row["银行原文件行号"])
    assert "2" in str(row["日记账原文件行号"])
    assert float(row["银行金额"]) == 120000
    assert float(row["日记账金额"]) == 120000
    assert row["风险等级"] == "高风险"
    assert row["系统结论"] == "疑点事项"


def test_精确匹配率与自动处理率分开且数值不超过百分之百():
    tables = _three_vs_five_reporter().build_report_tables(
        MatcherConfig()
    )
    summary = dict(
        zip(tables["核对结论"]["项目"], tables["核对结论"]["数值"])
    )

    assert isinstance(summary["逐笔精确匹配率"], (int, float))
    assert isinstance(summary["自动完成率"], (int, float))
    assert 0 <= summary["逐笔精确匹配率"] <= 1
    assert 0 <= summary["自动完成率"] <= 1
    assert {
        "自动确认组数",
        "整组勾稽组数",
        "自动归集事项数",
        "疑点事项数",
        "银行侧待查笔数",
        "日记账侧待查笔数",
    }.issubset(summary)


def test_首页同时汇总数量金额并给出三项建议动作():
    tables = _three_vs_five_reporter().build_report_tables(
        MatcherConfig()
    )
    summary = dict(
        zip(tables["核对结论"]["项目"], tables["核对结论"]["数值"])
    )

    assert {
        "自动确认金额",
        "整组勾稽金额",
        "自动归集金额",
        "疑点事项金额",
        "高风险事项金额",
        "建议动作1",
        "建议动作2",
        "建议动作3",
    }.issubset(summary)
    assert summary["整组勾稽金额"] == 500.0


def test_匹配组报告披露双方收支净额差异池和模型参与情况():
    groups = _three_vs_five_reporter().build_report_tables(
        MatcherConfig()
    )["整组勾稽"]

    assert {
        "银行收入",
        "银行支出",
        "银行净额",
        "日记账收入",
        "日记账支出",
        "日记账净额",
        "差异池ID",
        "是否使用大模型",
    }.issubset(groups.columns)


def test_空数据仍生成所有固定工作表和运行参数(tmp_path):
    empty = _report_df("2026-01-01", [], "bank")
    matcher = Matcher(empty, empty, MatcherConfig())
    reporter = Reporter(matcher)
    output = tmp_path / "empty.xlsx"

    reporter.generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output)

    required = {
        "核对结论",
        "疑点事项",
        "自动归集事项",
        "银行侧待查",
        "日记账侧待查",
        "逐笔匹配",
        "整组勾稽",
        "匹配组成",
        "每日统计",
        "月度统计",
        "运行参数",
    }
    assert required.issubset(workbook.sheetnames)


def test_用户文本不能被Excel当作公式执行(tmp_path):
    dangerous = ["=1+1", "+CMD", "-2+3", "@SUM(A1)", "\t=2+2"]
    bank = _report_df(
        "2026-01-01",
        [1, 2, 3, 4, 5],
        "bank",
        dangerous,
    )
    journal = _report_df("2026-01-01", [], "journal")
    reporter = Reporter(Matcher(bank, journal, MatcherConfig()))
    output = tmp_path / "safe.xlsx"

    reporter.generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output, data_only=False)

    sheet = workbook["银行侧待查"]
    summary_column = next(
        cell.column
        for cell in sheet[1]
        if cell.value == "摘要"
    )
    for row_number in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row_number, column=summary_column)
        assert cell.data_type != "f"


def test_外来表头不能被Excel当作公式执行(tmp_path):
    bank = _report_df("2026-01-01", [100], "bank", ["未匹配"])
    journal = _report_df("2026-01-01", [], "journal")
    raw_bank = pd.DataFrame({"=1+1": ["外来列"], "日期": ["2026-01-01"]})
    output = tmp_path / "safe-header.xlsx"

    Reporter(
        Matcher(bank, journal, MatcherConfig()),
        raw_bank=raw_bank,
    ).generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output, data_only=False)
    headers = list(workbook["银行侧待查"][1])

    assert all(cell.data_type != "f" for cell in headers)
    assert "'=1+1" in {cell.value for cell in headers}


def test_可选后续状态用公式联动首页处理进度(tmp_path):
    output = tmp_path / "progress.xlsx"
    _pending_reporter().generate_report(str(output), config=MatcherConfig())
    workbook = load_workbook(output, data_only=False)
    sheet = workbook["核对结论"]
    headers = {cell.value: cell.column for cell in sheet[1]}
    item_column = headers["项目"]
    value_column = headers["数值"]
    rows = {
        sheet.cell(row=row, column=item_column).value: sheet.cell(
            row=row,
            column=value_column,
        )
        for row in range(2, sheet.max_row + 1)
    }

    for item in ("待处理事项数", "已关注事项数", "已处理事项数", "无需处理事项数"):
        assert rows[item].data_type == "f"
        assert "疑点事项" in rows[item].value


def test_输入检查排在统计表之前且运行参数位于技术表区():
    reporter = _three_vs_five_reporter()
    reporter.precheck_report = InputPrecheckReport(
        items=(
            PrecheckItem(
                name="文件读取",
                bank_result="可用",
                journal_result="可用",
                comparison="可核对",
                status="通过",
                explanation="双方数据可计算",
            ),
        )
    )

    names = list(reporter.build_report_tables(MatcherConfig()))

    assert names.index("输入检查") < names.index("月度统计")
    assert names.index("每日统计") < names.index("运行参数")


def test_大模型明细字段完整且敏感内容被清理(tmp_path):
    matcher = _pending_reporter().matcher
    candidate = matcher.selected_candidates[0]
    matcher.llm_records = [
        LLMDecisionRecord(
            request_id="R1",
            candidate_ids=(candidate.candidate_id,),
            sent_fields=("日期", "金额", "摘要"),
            selected_candidate_id=candidate.candidate_id,
            semantic_score=88,
            reason="账号622233334444支持匹配",
            supporting_evidence=("摘要一致",),
            conflicting_evidence=(),
            uncertainty="低",
            suggested_status="自动确认",
            provider="在线 API",
            model="mock-model",
            started_at="2026-01-01T10:00:00+08:00",
            duration_ms=120,
            usage={"input_tokens": 10},
            raw_response=(
                "账号622233334444；"
                "https://example.com/v1?api_key=secret-key"
            ),
        )
    ]
    output = tmp_path / "llm.xlsx"

    Reporter(matcher).generate_report(
        str(output),
        config=MatcherConfig(),
    )
    table = pd.read_excel(output, sheet_name="大模型辅助明细")
    body = table.to_string()

    assert {
        "本地文字分",
        "候选ID",
        "是否模型选择",
        "银行原文件行号",
        "日记账原文件行号",
        "银行日期",
        "日记账日期",
        "银行金额",
        "日记账金额",
        "实际发送字段",
        "接口地址",
        "本地综合可信度",
        "模型语义分",
        "判断理由",
        "支持证据",
        "冲突证据",
        "不确定性",
        "最终状态",
        "服务",
        "模型",
        "耗时毫秒",
        "用量",
        "是否降级",
        "实际执行重要性水平",
        "明显微小错报临界值",
        "自动确认最低综合可信度",
        "脱敏原始回答",
    }.issubset(table.columns)
    assert str(candidate.bank_idxs[0] + 2) in str(
        table.iloc[0]["银行原文件行号"]
    )
    assert "622233334444" not in body
    assert "secret-key" not in body
