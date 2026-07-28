"""不依赖图形界面的完整核对流程编排。"""

from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig
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


def run_reconciliation(
    bank_path: str,
    journal_path: str,
    bank_mapping: dict[str, object],
    journal_mapping: dict[str, object],
    matcher_config: MatcherConfig,
    llm_config: Optional[LLMConfig] = None,
    logger: Optional[Callable[[str], None]] = None,
    *,
    bank_skiprows: int = 0,
    journal_skiprows: int = 0,
    date_format: str = "auto",
    output_path: Optional[str | Path] = None,
    progress_callback: Optional[Callable[[float], None]] = None,
    matcher_ready: Optional[Callable[[Matcher], None]] = None,
) -> Path:
    """读取、标准化、匹配并生成 Excel 报告。

    这个入口不依赖 Tk，可由图形界面、自动测试或批处理共同调用。
    进度统一使用 0 到 1，便于界面直接显示百分比。
    """

    log = logger or (lambda _message: None)
    collector = ParseErrorCollector()
    loader = DataLoader(logger=log, error_collector=collector)

    if progress_callback:
        progress_callback(0.1)
    raw_bank = loader.load_file(bank_path, skiprows=bank_skiprows)
    raw_journal = loader.load_file(
        journal_path,
        skiprows=journal_skiprows,
    )
    bank = loader.standardize_data(
        raw_bank.copy(),
        bank_mapping,
        "bank",
        date_format,
        skiprows_offset=bank_skiprows,
    )
    journal = loader.standardize_data(
        raw_journal.copy(),
        journal_mapping,
        "journal",
        date_format,
        skiprows_offset=journal_skiprows,
    )

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
