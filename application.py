"""不依赖图形界面的完整核对流程编排。"""

from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig
from input_precheck import (
    InputPrecheckBlockedError,
    InputPrecheckReport,
    PrecheckItem,
    TableStructure,
    build_input_precheck,
    derive_header_columns,
)
from llm_assistant import LLMConfig, LLMAssistant
from matcher import Matcher
from reporter import Reporter


def _default_output_path(bank_path: str, journal_path: str) -> Path:
    bank = Path(bank_path)
    journal = Path(journal_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return bank.parent / (
        f"{bank.stem}_vs_{journal.stem}_核对报告_{timestamp}.xlsx"
    )


def _blocked_report(
    name: str,
    explanation: str,
) -> InputPrecheckReport:
    return InputPrecheckReport(
        items=(
            PrecheckItem(
                name=name,
                bank_result="未通过",
                journal_result="未通过",
                comparison="无法进入正式匹配",
                status="阻止",
                explanation=explanation,
            ),
        ),
    )


def _adopt_structure(
    path: str,
    detected: TableStructure,
    skiprows: Optional[int],
    header_rows: Optional[int],
) -> TableStructure:
    adopted_skiprows = detected.skiprows if skiprows is None else int(skiprows)
    adopted_header_rows = detected.header_rows if header_rows is None else int(header_rows)
    if (adopted_skiprows, adopted_header_rows) == (
        detected.skiprows,
        detected.header_rows,
    ):
        return detected
    columns = derive_header_columns(path, adopted_skiprows, adopted_header_rows)
    return TableStructure(
        skiprows=adopted_skiprows,
        header_rows=adopted_header_rows,
        columns=columns,
        score=detected.score,
        ambiguous=detected.ambiguous,
        candidates=detected.candidates,
        explanation=(
            f"采用用户设置：第{adopted_skiprows + 1}行起、"
            f"共{adopted_header_rows}行表头"
        ),
    )


def run_reconciliation(
    bank_path: str,
    journal_path: str,
    bank_mapping: dict[str, object],
    journal_mapping: dict[str, object],
    matcher_config: MatcherConfig,
    llm_config: Optional[LLMConfig] = None,
    logger: Optional[Callable[[str], None]] = None,
    *,
    bank_skiprows: Optional[int] = None,
    journal_skiprows: Optional[int] = None,
    bank_header_rows: Optional[int] = None,
    journal_header_rows: Optional[int] = None,
    date_format: str = "auto",
    output_path: Optional[str | Path] = None,
    progress_callback: Optional[Callable[[float], None]] = None,
    matcher_ready: Optional[Callable[[Matcher], None]] = None,
    precheck_warning_callback: Optional[
        Callable[[InputPrecheckReport], bool]
    ] = None,
) -> Path:
    """读取、标准化、匹配并生成 Excel 报告。

    这个入口不依赖 Tk，可由图形界面、自动测试或批处理共同调用。
    进度统一使用 0 到 1，便于界面直接显示百分比。
    """

    log = logger or (lambda _message: None)
    collector = ParseErrorCollector()
    loader = DataLoader(logger=log, error_collector=collector)

    bank_structure: Optional[TableStructure] = None
    journal_structure: Optional[TableStructure] = None
    try:
        bank_structure = _adopt_structure(
            bank_path,
            loader.detect_table_structure(bank_path),
            bank_skiprows,
            bank_header_rows,
        )
        journal_structure = _adopt_structure(
            journal_path,
            loader.detect_table_structure(journal_path),
            journal_skiprows,
            journal_header_rows,
        )
        raw_bank = loader.load_file(
            bank_path,
            skiprows=bank_structure.skiprows,
            header_rows=bank_structure.header_rows,
            derived_columns=bank_structure.columns,
        )
        raw_journal = loader.load_file(
            journal_path,
            skiprows=journal_structure.skiprows,
            header_rows=journal_structure.header_rows,
            derived_columns=journal_structure.columns,
        )
    except Exception as exc:
        raise InputPrecheckBlockedError(
            _blocked_report(
                "文件读取",
                f"文件或表头无法正常读取：{exc}",
            )
        ) from exc

    if progress_callback:
        progress_callback(0.1)
    try:
        bank = loader.standardize_data(
            raw_bank.copy(),
            bank_mapping,
            "bank",
            date_format,
            skiprows_offset=bank_structure.skiprows,
            header_rows=bank_structure.header_rows,
        )
        journal = loader.standardize_data(
            raw_journal.copy(),
            journal_mapping,
            "journal",
            date_format,
            skiprows_offset=journal_structure.skiprows,
            header_rows=journal_structure.header_rows,
        )
    except Exception as exc:
        errors = collector.get_all_errors()
        has_direction_error = any(
            error.get("type") == "方向解析失败" for error in errors
        )
        name = "金额方向" if has_direction_error else "必填字段或数据解析"
        explanation = (
            "方向列存在无法识别的值，收入和支出方向不可靠。"
            if has_direction_error
            else f"输入列或数据无法形成有效交易：{exc}"
        )
        raise InputPrecheckBlockedError(
            _blocked_report(
                name,
                explanation,
            )
        ) from exc

    error_summary = collector.get_summary()
    if error_summary["总计"] > 0:
        log(
            "解析异常统计："
            f"金额失败{error_summary['金额解析失败']}条，"
            f"方向失败{error_summary['方向解析失败']}条，"
            f"日期失败{error_summary['日期解析失败']}条，"
            f"汇总行丢弃{error_summary['被丢弃的汇总行']}条，"
            f"空日期{error_summary['空日期行']}条"
        )

    precheck_report = build_input_precheck(
        raw_bank=raw_bank,
        raw_journal=raw_journal,
        bank=bank,
        journal=journal,
        bank_mapping=bank_mapping,
        journal_mapping=journal_mapping,
        bank_structure=bank_structure,
        journal_structure=journal_structure,
        parse_errors=collector.get_all_errors(),
    )
    if precheck_report.has_blockers:
        raise InputPrecheckBlockedError(precheck_report)
    if precheck_report.has_warnings:
        log("输入预检查提示：\n" + precheck_report.warning_message())
        if (
            precheck_warning_callback is not None
            and not precheck_warning_callback(precheck_report)
        ):
            raise InterruptedError("用户返回调整输入")
    else:
        log("输入预检查通过")
    if progress_callback:
        progress_callback(0.3)

    effective_llm_config = llm_config or LLMConfig()
    assistant = (
        LLMAssistant(effective_llm_config)
        if effective_llm_config.enabled
        else None
    )

    def report_matcher_progress(value: float) -> None:
        if not progress_callback:
            return
        normalized = value / 100 if value > 1 else value
        progress_callback(max(0.3, min(0.75, normalized)))

    matcher = Matcher(
        bank,
        journal,
        matcher_config,
        logger=log,
        progress_callback=report_matcher_progress,
        llm_assistant=assistant,
    )
    if matcher_ready:
        matcher_ready(matcher)
    matcher.run()
    if matcher.stopping:
        raise InterruptedError("核对任务已取消")
    if progress_callback:
        progress_callback(0.8)

    reporter = Reporter(
        matcher,
        raw_bank=raw_bank,
        raw_journal=raw_journal,
        bank_mapping=bank_mapping,
        journal_mapping=journal_mapping,
        logger=log,
        error_collector=collector,
        precheck_report=precheck_report,
    )
    destination = (
        Path(output_path)
        if output_path is not None
        else _default_output_path(bank_path, journal_path)
    )
    destination.parent.mkdir(parents=True, exist_ok=True)
    reporter.generate_report(
        str(destination),
        config=matcher_config,
        bank_path=bank_path,
        journal_path=journal_path,
        date_format=date_format,
    )
    if progress_callback:
        progress_callback(1.0)
    return destination
