"""不依赖图形界面的完整核对流程编排。"""

import hashlib
import json
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig
from input_precheck import (
    InputPrecheckBlockedError,
    InputPrecheckReport,
    PrecheckItem,
    TableStructure,
    build_input_precheck,
    derive_header_columns,
    write_input_problem_report,
)
from llm_assistant import LLMConfig, LLMAssistant
from matcher import Matcher
from reporter import Reporter


def _map_matcher_progress(value: float) -> float:
    """把匹配器的 0—100 进度线性映射到整项任务的 30%—75%。"""
    normalized = float(value)
    if normalized > 1:
        normalized /= 100
    normalized = max(0.0, min(1.0, normalized))
    return 0.3 + normalized * 0.45


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
                status="无法计算",
                explanation=explanation,
            ),
        ),
    )


def _source_info(
    path_value: str,
    source_name: str,
    mapping: dict[str, object],
    structure: Optional[TableStructure] = None,
) -> dict[str, object]:
    path = Path(path_value)
    digest = ""
    size: int | str = ""
    sheet = ""
    if path.is_file():
        hasher = hashlib.sha256()
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                hasher.update(chunk)
        digest = hasher.hexdigest()
        size = path.stat().st_size
        if path.suffix.lower() == ".csv":
            sheet = "CSV"
        else:
            try:
                with pd.ExcelFile(path) as workbook:
                    sheet = workbook.sheet_names[0] if workbook.sheet_names else ""
            except Exception:
                sheet = "无法读取"
    return {
        "来源": source_name,
        "文件路径": str(path.resolve(strict=False)),
        "文件SHA256": digest,
        "文件大小": size,
        "工作表": sheet,
        "表头起始行": "" if structure is None else structure.skiprows + 1,
        "表头行数": "" if structure is None else structure.header_rows,
        "采用映射": json.dumps(mapping, ensure_ascii=False, sort_keys=True, default=str),
    }


def _problem_report_path(destination: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    return destination.with_name(f"{destination.stem}_输入问题_{timestamp}.xlsx")


def _blocked_error_with_report(
    report: InputPrecheckReport,
    destination: Path,
) -> InputPrecheckBlockedError:
    problem_path = _problem_report_path(destination)
    try:
        written = write_input_problem_report(report, problem_path)
        return InputPrecheckBlockedError(report, report_path=written)
    except Exception as exc:
        return InputPrecheckBlockedError(
            report,
            report_write_error=f"输入问题报告写入失败：{exc}",
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
    destination = (
        Path(output_path)
        if output_path is not None
        else _default_output_path(bank_path, journal_path)
    )
    total_started = time.perf_counter()
    if progress_callback:
        progress_callback(0.0)

    bank_structure: Optional[TableStructure] = None
    journal_structure: Optional[TableStructure] = None
    load_started = time.perf_counter()
    log("开始读取银行流水和银行存款序时账")
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
        log(
            f"文件读取完成：银行流水 {len(raw_bank):,} 行，"
            f"银行存款序时账 {len(raw_journal):,} 行，"
            f"耗时 {time.perf_counter() - load_started:.1f} 秒"
        )
    except Exception as exc:
        report = _blocked_report(
                "文件读取",
                f"文件或表头无法正常读取：{exc}",
            )
        report = InputPrecheckReport(
            items=report.items,
            source_info=(
                _source_info(bank_path, "银行流水", bank_mapping, bank_structure),
                _source_info(journal_path, "银行日记账", journal_mapping, journal_structure),
            ),
        )
        raise _blocked_error_with_report(report, destination) from exc

    if progress_callback:
        progress_callback(0.1)
    bank = pd.DataFrame()
    journal = pd.DataFrame()
    standardize_started = time.perf_counter()
    log("开始解析日期、金额方向和审计辅助字段")
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
        log(
            f"数据解析完成：银行流水有效交易 {len(bank):,} 行，"
            f"银行存款序时账有效交易 {len(journal):,} 行，"
            f"耗时 {time.perf_counter() - standardize_started:.1f} 秒"
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
        report = _blocked_report(
                name,
                explanation,
            )
        evidence = build_input_precheck(
            raw_bank=raw_bank, raw_journal=raw_journal,
            bank=bank, journal=journal,
            bank_mapping=bank_mapping, journal_mapping=journal_mapping,
            bank_structure=bank_structure, journal_structure=journal_structure,
            parse_errors=errors,
        )
        report = InputPrecheckReport(
            items=report.items,
            source_info=(
                _source_info(bank_path, "银行流水", bank_mapping, bank_structure),
                _source_info(journal_path, "银行日记账", journal_mapping, journal_structure),
            ),
            population_rows=evidence.population_rows,
            parse_errors=tuple(errors),
        )
        raise _blocked_error_with_report(report, destination) from exc

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

    precheck_started = time.perf_counter()
    log("开始执行账户、币种、期间、金额和数据人口检查")
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
        source_info=(
            _source_info(bank_path, "银行流水", bank_mapping, bank_structure),
            _source_info(journal_path, "银行日记账", journal_mapping, journal_structure),
        ),
    )
    if precheck_report.has_blockers:
        raise _blocked_error_with_report(precheck_report, destination)
    if precheck_report.has_warnings:
        log(
            "输入预检查发现疑点，程序已自动继续并将在报告中披露：\n"
            + precheck_report.warning_message()
        )
    else:
        log("输入预检查通过")
    log(f"输入预检查完成，耗时 {time.perf_counter() - precheck_started:.1f} 秒")
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
        progress_callback(_map_matcher_progress(value))

    matcher = Matcher(
        bank,
        journal,
        matcher_config,
        logger=log,
        progress_callback=report_matcher_progress,
        llm_assistant=assistant,
        overall_control=precheck_report.overall_control,
    )
    if matcher_ready:
        matcher_ready(matcher)
    match_started = time.perf_counter()
    log(
        f"开始形成核对候选：银行流水 {len(bank):,} 行，"
        f"银行存款序时账 {len(journal):,} 行"
    )
    matcher.run()
    if matcher.stopping:
        raise InterruptedError("核对任务已取消")
    if progress_callback:
        progress_callback(0.8)
    log(
        f"匹配完成：选中关系 {len(matcher.selected_candidates):,} 组，"
        f"耗时 {time.perf_counter() - match_started:.1f} 秒"
    )

    report_started = time.perf_counter()
    log("开始生成和保存审计核对报告")
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
    log(
        f"全部完成：报告生成耗时 {time.perf_counter() - report_started:.1f} 秒，"
        f"总耗时 {time.perf_counter() - total_started:.1f} 秒"
    )
    return destination
