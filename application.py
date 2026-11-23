"""不依赖图形界面的完整核对流程编排。"""

from collections import Counter
import hashlib
import json
import re
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
    statement_filename_identity,
    _mapping_problems,
)
from llm_assistant import LLMConfig, LLMAssistant
from matcher import Matcher
from reporter import Reporter
from 项目记录 import LocalProjectStore
from 列识别 import ColumnRecognitionError, normalized_identity, recognize_columns


def _balance_scope_identity(bank_path, journal_path, raw_journal, control, *,
                            verified_account="", stated_accounts=()):
    """先核实账户，再核实双方收支、期初期末和余额连续性。"""
    if any(getattr(control, field) != "通过" for field in (
        "amount_status", "bank_balance_status", "journal_balance_status",
    )) or control.initial_balance_diff != 0 or control.ending_balance_diff != 0:
        return None
    allowed_reasons = ("核对账户未完全验证", "核对币种未完全验证", "金额口径未完全验证")
    if verified_account:
        allowed_reasons += ("双方起止日期不一致",)
    if any(not reason.startswith(allowed_reasons) for reason in control.reasons) or control.continuity_anomalies:
        return None

    file_identity = statement_filename_identity(Path(bank_path))
    journal_identity = statement_filename_identity(Path(journal_path))
    if verified_account and control.period_status in {"通过", "疑点"}:
        if file_identity and file_identity == journal_identity and file_identity[-1] == verified_account:
            return file_identity
        return "", "本账户", verified_account
    if (control.period_status != "通过" or not file_identity or file_identity != journal_identity
            or not re.fullmatch(r"[0-9]{10,30}", file_identity[-1])
            or any(account != file_identity[-1] for account in stated_accounts)
            or "辅助核算名称" not in raw_journal):
        return None
    names = raw_journal["辅助核算名称"].fillna("").astype(str).str.strip()
    if names.empty or not names.str.extract(r"([0-9]+)$", expand=False).eq(file_identity[-1]).all():
        return None
    return file_identity


def _balanced_posting_days(frame):
    """按原表顺序验证日内余额链，按有方向的金额统计逐笔组成。"""
    days = {}
    for day, rows in frame.groupby(pd.to_datetime(frame["date"]).dt.normalize(), sort=False):
        if len(rows) < 2 or rows[["amount", "balance"]].isna().any().any():
            continue
        rows = rows.sort_values("original_file_row", kind="stable")
        before = rows["balance"] - rows["amount"]
        if before.iloc[1:].tolist() != rows["balance"].iloc[:-1].tolist():
            continue
        amounts = Counter(rows["amount"])
        unique = {amount: index for index, amount in rows["amount"].items() if amounts[amount] == 1 and amount}
        days[day] = ((before.iloc[0], rows["balance"].iloc[-1], amounts), unique)
    return days


def _verified_balance_pairs(bank_path, journal_path, raw_bank, raw_journal, bank, journal, control, *,
                            bank_first_row=2, verified_account="", stated_accounts=()):
    identity = _balance_scope_identity(
        bank_path, journal_path, raw_journal, control,
        verified_account=verified_account, stated_accounts=stated_accounts,
    )
    if identity is None:
        return {}
    company, own_bank, account = identity

    def fingerprints(frame):
        rows = frame.dropna(subset=["date", "amount", "balance"])
        return rows.assign(day=pd.to_datetime(rows["date"]).dt.normalize()).groupby(["day", "amount", "balance"], sort=False).groups

    journal_keys = fingerprints(journal)
    pairs = {}
    for key, bank_indexes in fingerprints(bank).items():
        journal_indexes = journal_keys.get(key, ())
        if len(bank_indexes) != 1 or len(journal_indexes) != 1 or not key[1]:
            continue
        b, j = bank_indexes[0], journal_indexes[0]
        pairs[b, j] = (key[0], key[1], False)

    journal_days = _balanced_posting_days(journal)
    for day, (bank_controls, bank_unique) in _balanced_posting_days(bank).items():
        if day not in journal_days or bank_controls != journal_days[day][0]:
            continue
        for amount, b in bank_unique.items():
            j = journal_days[day][1][amount]
            pairs.setdefault((b, j), (day, amount, True))

    row_numbers = raw_bank.get("__file_row__", raw_bank.index + bank_first_row)
    bank_rows = raw_bank.set_axis(row_numbers)
    verified = {}
    for (b, j), (day, amount, different_order) in pairs.items():
        bank_original = bank_rows.loc[int(bank.at[b, "original_file_row"])]
        other_account = str(bank_original.get("对方账号", "")).strip()
        voucher_value = journal.at[j, "voucher_no"]
        voucher = str(voucher_value).strip() if pd.notna(voucher_value) else ""
        summary = str(journal.at[j, "summary"]).strip()
        bank_party = str(bank_original.get("对方户名", "")).strip()
        if (company and bank_party == company and re.fullmatch(r"[0-9]{10,30}", other_account)
            and other_account != account and voucher and summary in {"转账", "内部转账", "同名账户划款"}):
            other_bank = bank_original.get("对方开户行")
            other_label = f"{other_bank if pd.notna(other_bank) else '对方账户'}尾号{other_account[-4:]}"
            own_label = f"{own_bank}尾号{account[-4:]}"
            start, end = (other_label, own_label) if amount > 0 else (own_label, other_label)
            detail = f"{day:%Y-%m-%d}，{company}从{start}转入{end} {abs(amount):,.2f} 元。"
        else:
            action = "收款" if amount > 0 else "付款"
            detail = f"{day:%Y-%m-%d}，{own_bank}尾号{account[-4:]}{action} {abs(amount):,.2f} 元。"
        ledger = f"第{voucher}号凭证" if voucher else "账上记录"
        if different_order:
            verified[int(b), int(j)] = (
                detail + f"{ledger}有唯一同日同额记录。双方当天流水顺序不同，"
                "逐笔收付金额组成一致，日初、日末余额也一致。"
            )
        else:
            verified[int(b), int(j)] = detail + f"{ledger}的日期、金额和交易后余额与银行一致。"
    return verified


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
                from 直读XML import 读表名
                表名 = 读表名(path)
                sheet = 表名[0] if 表名 else ""
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
    bank_mapping: Optional[dict[str, object]],
    journal_mapping: Optional[dict[str, object]],
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
    project_store: Optional[LocalProjectStore] = None,
    cancel_requested: Optional[Callable[[], bool]] = None,
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
    for source_path in (Path(bank_path), Path(journal_path)):
        if destination.resolve() == source_path.resolve() or (
            destination.exists() and source_path.exists()
            and destination.samefile(source_path)
        ):
            raise ValueError("报告不能覆盖原始输入文件，请另存为新文件")
    if destination.exists():
        raise FileExistsError("报告文件已存在，请另存为新文件，避免覆盖人工核对结果")
    matcher = None

    def check_cancelled() -> None:
        if (cancel_requested and cancel_requested()) or (matcher is not None and matcher.stopping):
            raise InterruptedError("核对任务已取消")

    total_started = time.perf_counter()
    if progress_callback:
        progress_callback(0.0)

    bank_structure: Optional[TableStructure] = None
    journal_structure: Optional[TableStructure] = None
    load_started = time.perf_counter()
    log("开始读取银行流水和银行存款序时账")
    check_cancelled()
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

    column_issues = []
    mappings = []
    for source, path, raw, supplied in (
        ("bank", bank_path, raw_bank, bank_mapping),
        ("journal", journal_path, raw_journal, journal_mapping),
    ):
        choices = {}
        if not supplied or supplied.get("_column_auto"):
            result = recognize_columns(
                raw, source, overrides=(supplied or {}).get("_column_overrides"), date_format=date_format,
            )
            mapping, choices, problems = result.mapping, result.choices, result.problems
            if not choices and not problems:
                problems = _mapping_problems(mapping, raw.columns, source)
            auxiliary = (supplied or {}).get("auxiliary_text_columns")
            if auxiliary is not None:
                mapping["auxiliary_text_columns"] = [column for column in auxiliary if column in raw.columns]
        else:
            mapping = supplied
            problems = _mapping_problems(mapping, raw.columns, source)
        if problems or choices:
            column_issues.append({
                "侧": source, "来源": "银行流水" if source == "bank" else "日记账",
                "文件": Path(path).name, "候选列": choices, "问题": problems,
            })
        mappings.append(mapping)
    bank_mapping, journal_mapping = mappings
    if column_issues:
        raise ColumnRecognitionError(column_issues)

    check_cancelled()
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

    check_cancelled()
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
    log("开始执行账户、币种、期间、金额和数据入口检查")
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
    check_cancelled()

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

    account_passed = any(item.name == "核对账户" and item.status == "通过" for item in precheck_report.items)
    stated_accounts = {
        normalized_identity(value, "account")
        for frame in (bank, journal) for value in frame["account"].dropna().unique()
    } - {""}
    verified_account = ""
    if account_passed and len(stated_accounts) == 1:
        account = next(iter(stated_accounts))
        if re.fullmatch(r"[0-9A-Z]+", account):
            verified_account = account

    matcher = Matcher(
        bank,
        journal,
        matcher_config,
        logger=log,
        progress_callback=report_matcher_progress,
        llm_assistant=assistant,
        overall_control=precheck_report.overall_control,
        balance_scope_verified=_balance_scope_identity(
            bank_path, journal_path, raw_journal, precheck_report.overall_control,
            verified_account=verified_account, stated_accounts=stated_accounts,
        ) is not None,
        verified_balance_pairs=_verified_balance_pairs(
            bank_path, journal_path, raw_bank, raw_journal, bank, journal, precheck_report.overall_control,
            bank_first_row=bank_structure.skiprows + bank_structure.header_rows + 1,
            verified_account=verified_account, stated_accounts=stated_accounts,
        ),
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
    check_cancelled()
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
        cancel_check=check_cancelled,
    )
    report_elapsed = time.perf_counter() - report_started
    history_started = time.perf_counter()
    log("开始保存列映射模板和核对项目历史")
    try:
        store = project_store or LocalProjectStore()
        store.save_mapping_template(
            "bank",
            raw_bank.columns,
            bank_mapping,
            f"银行流水-{Path(bank_path).stem}",
        )
        store.save_mapping_template(
            "journal",
            raw_journal.columns,
            journal_mapping,
            f"银行存款序时账-{Path(journal_path).stem}",
        )
        store.save_project(
            bank_path=bank_path,
            journal_path=journal_path,
            report_path=destination,
            bank_mapping=bank_mapping,
            journal_mapping=journal_mapping,
            parameters=matcher_config,
            result_counts={
                "bank_rows": len(bank),
                "journal_rows": len(journal),
                "selected_groups": len(matcher.selected_candidates),
                "business_events": len(matcher.business_events),
                "business_clues": len(matcher.business_clues),
            },
        )
        log(f"项目历史保存完成，耗时 {time.perf_counter() - history_started:.1f} 秒")
    except (OSError, ValueError, TypeError) as exc:
        log(f"报告已保存至：{destination}；列映射模板或项目历史未保存：{exc}")
    if progress_callback:
        progress_callback(1.0)
    log(
        f"全部完成：报告生成耗时 {report_elapsed:.1f} 秒，"
        f"总耗时 {time.perf_counter() - total_started:.1f} 秒"
    )
    return destination
