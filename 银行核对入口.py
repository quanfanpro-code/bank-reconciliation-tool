"""为 Agent 提供单户及目录核对入口。用户设置由 Agent 在问答中收集。"""

import argparse
from collections import Counter, defaultdict
from dataclasses import asdict, replace
from datetime import datetime
from decimal import Decimal
import json
from pathlib import Path
import time
from uuid import uuid4

from data_structures import MatcherConfig
from input_precheck import statement_filename_identity
from 列识别 import ColumnRecognitionError, FIELD_LABELS


def pair_inputs(bank, journal):
    if bank.is_file() and journal.is_file():
        if bank.resolve() == journal.resolve():
            raise ValueError("银行流水和日记账不能使用同一个文件")
        left, right = statement_filename_identity(bank), statement_filename_identity(journal)
        if left and right and left != right:
            raise ValueError("两个文件名标明的公司、银行或账号不一致")
        return [(bank, journal)], []
    if not (bank.is_dir() and journal.is_dir()):
        raise ValueError("请提供两个文件，或分别提供银行流水目录和日记账目录")
    if bank.resolve() == journal.resolve():
        raise ValueError("银行流水和日记账应分别指定目录")

    files = {"银行流水": defaultdict(list), "日记账": defaultdict(list)}
    problems = []
    for label, directory in (("银行流水", bank), ("日记账", journal)):
        for path in sorted(directory.iterdir()):
            if not path.is_file() or path.suffix.lower() not in {".xlsx", ".csv"}:
                continue
            if path.name.startswith(("~$", ".")) or any(word in path.stem for word in ("核对报告", "导入模板", "输入问题")):
                continue
            identity = statement_filename_identity(path)
            if identity is None:
                problems.append({"原因": f"{label}文件名未标明公司、银行和账号", "文件": str(path)})
            else:
                files[label][identity].append(path)

    pairs = []
    for identity in sorted(files["银行流水"].keys() | files["日记账"].keys()):
        left, right = (files[label].get(identity, []) for label in files)
        if len(left) == len(right) == 1:
            pairs.append((left[0], right[0]))
            continue
        reasons = [f"缺少{label}" if not entries else f"{label}有多个文件"
                   for label, entries in zip(files, (left, right)) if len(entries) != 1]
        problems.append({"公司": identity[0], "银行": identity[1], "账号": identity[2], "原因": "；".join(reasons),
                         "银行流水": list(map(str, left)), "日记账": list(map(str, right))})
    return pairs, problems


def _review_counts():
    return Counter({"已完成": 0, "超时": 0, "未复核": 0})


class AgentReview:
    """在本次任务目录交接候选和 Agent 答复。"""

    def __init__(self, directory, timeout=120, sources=()):
        self.directory = directory
        self.timeout = timeout
        self.sources = sources
        self.review_counts = _review_counts()
        directory.mkdir(parents=True, exist_ok=True)

    def evaluate_candidates(self, request):
        from data_structures import LLMDecisionRecord
        from llm_assistant import ONLINE_ALLOWED_FIELDS

        ids = tuple(candidate.candidate_id for candidate in request.candidates)
        verdict = LLMDecisionRecord(request.request_id, ids, provider="当前 Agent", fallback_used=True)
        if len(ids) < 2:
            self.review_counts["未复核"] += 1
            return verdict

        token = uuid4().hex
        candidates = []
        for candidate in request.candidates:
            summary = asdict(candidate)
            for key in ("bank_fields", "journal_fields"):
                summary[key] = {name: value for name, value in summary[key].items() if name in ONLINE_ALLOWED_FIELDS}
            candidates.append(summary)
        answer, error = self.exchange({"candidates": candidates}, token=token)
        if answer is None:
            verdict.error = error
            return verdict
        selected = answer.get("selected_candidate_id")
        score = answer.get("semantic_score")
        reason = answer.get("reason")
        if selected in ids and type(score) is int and 0 <= score <= 100 and isinstance(reason, str) and reason.strip():
            evidence = answer.get("source_evidence")
            fields = ("bank_field", "journal_field", "quote")
            source_evidence = None
            if isinstance(evidence, dict) and all(
                isinstance(evidence.get(field), str) and evidence[field].strip() for field in fields
            ):
                source_evidence = tuple(evidence[field].strip() for field in fields)
            self.review_counts["已完成"] += 1
            return replace(verdict, selected_candidate_id=selected, semantic_score=score,
                           reason=reason.strip(), source_evidence=source_evidence, fallback_used=False)
        if selected == "":
            self.review_counts["已完成"] += 1
            verdict.error = "当前 Agent 保留疑点"
        else:
            self.review_counts["未复核"] += 1
            verdict.error = "Agent 返回了无效的候选或依据"
        return verdict


    def exchange(self, payload, *, token=None):
        """发送一个独立计时的任务并读取同一任务的答复。"""
        token = token or uuid4().hex
        request_path = self.directory / f"request-{token}.json"
        staged = request_path.with_suffix(".new")
        deadline = time.monotonic() + self.timeout
        expires_at = time.time() + self.timeout
        staged.write_text(json.dumps({
            **payload, "id": token, "sources": self.sources, "expires_at": expires_at,
        }, ensure_ascii=False), encoding="utf-8")
        staged.replace(request_path)
        response_path = self.directory / f"answer-{token}.json"
        while not response_path.is_file() and (remaining := deadline - time.monotonic()) > 0:
            time.sleep(min(.1, remaining))
        if not response_path.is_file() or response_path.stat().st_mtime > expires_at:
            self.review_counts["超时"] += 1
            return None, "当前 Agent 未在等待时间内返回答复"
        try:
            answer = json.loads(response_path.read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError, ValueError):
            answer = None
        if not isinstance(answer, dict) or answer.get("id") != token:
            self.review_counts["未复核"] += 1
            return None, "Agent 答复无法关联到本次请求"
        return answer, ""


def run_accounts(bank, journal, output, config, session, *, log=print,
                 bank_columns=None, journal_columns=None, matcher_factory=None, coverage_evidence=None):
    from application import run_reconciliation
    from llm_assistant import LLMConfig
    from 项目记录 import LocalProjectStore

    pairs, problems = pair_inputs(bank, journal)
    if (bank_columns or journal_columns) and not (bank.is_file() and journal.is_file()):
        raise ValueError("指定字段时请提供需要补跑的两份文件")
    output.mkdir(parents=True, exist_ok=True)
    suffix = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid4().hex[:8]
    store = LocalProjectStore(session / "项目记录")
    results = []
    totals = _review_counts()
    for bank_file, journal_file in pairs:
        reviewer = None
        try:
            target = output / f"{bank_file.stem}_核对报告_{suffix}.xlsx"
            reviewer = AgentReview(session / "复核", sources=(str(bank_file), str(journal_file)))

            def set_reviewer(matcher):
                matcher.llm_assistant = reviewer
                matcher.coverage_evidence = {
                    side: (coverage_evidence or {}).get(side, {}).get(source.name, {})
                    for side, source in (("bank", bank_file), ("journal", journal_file))
                }

            report = run_reconciliation(
                str(bank_file), str(journal_file),
                {"_column_auto": True, "_column_overrides": bank_columns},
                {"_column_auto": True, "_column_overrides": journal_columns},
                config, llm_config=LLMConfig(enabled=False), logger=log,
                output_path=target, project_store=store, matcher_ready=set_reviewer,
                matcher_factory=matcher_factory,
            )
            results.append({"银行流水": str(bank_file), "日记账": str(journal_file), "报告": str(report),
                            "Agent复核": dict(reviewer.review_counts)})
        except ColumnRecognitionError as exc:
            problems.append({
                "银行流水": str(bank_file), "日记账": str(journal_file),
                "原因": str(exc), "字段问题": exc.issues,
            })
        except Exception as exc:
            problems.append({"银行流水": str(bank_file), "日记账": str(journal_file), "原因": str(exc)})
        finally:
            if reviewer is not None:
                totals.update(reviewer.review_counts)
    return {"完成": results, "待处理": problems, "Agent复核": dict(totals)}


def _column_overrides(parser, entries):
    result = {}
    for entry in entries:
        role, separator, column = entry.partition("=")
        if not separator or role not in FIELD_LABELS or not column or role in result:
            parser.error(f"无效或重复的字段设置：{entry}")
        result[role] = column
    return result


def main(*, matcher_factory=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bank", required=True, type=Path)
    parser.add_argument("--journal", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--session", required=True, type=Path)
    if matcher_factory is not None:
        parser.add_argument("--source-evidence", type=Path, help="从已提供原件摘录的余额及逐笔查询范围")
    parser.add_argument("--trivial", required=True, type=Decimal, help="用户确认的明显微小错报临界值")
    parser.add_argument("--materiality", required=True, type=Decimal, help="用户确认的实际执行重要性水平")
    parser.add_argument("--tolerance-days", required=True, type=int, help="用户确认的日期容差天数，不改时为 31")
    parser.add_argument("--bank-column", action="append", default=[], metavar="字段=原列名")
    parser.add_argument("--journal-column", action="append", default=[], metavar="字段=原列名")
    args = parser.parse_args()
    if args.tolerance_days is not None and args.tolerance_days < 0:
        parser.error("日期容差必须为非负整数")
    bank_columns = _column_overrides(parser, args.bank_column)
    journal_columns = _column_overrides(parser, args.journal_column)
    settings = {
        "clearly_trivial_threshold": args.trivial,
        "performance_materiality": args.materiality,
        "tolerance_days": args.tolerance_days,
        "dfs_date_window": args.tolerance_days,
    }
    config = replace(MatcherConfig(), **{
        key: value for key, value in settings.items()
        if value is not None
    })
    coverage_evidence = None
    if getattr(args, "source_evidence", None):
        from 核对期间 import parse_coverage_evidence
        try:
            coverage_evidence = parse_coverage_evidence(json.loads(args.source_evidence.read_text(encoding="utf-8-sig")))
        except (OSError, TypeError, KeyError, ValueError) as exc:
            parser.error(f"来源凭据有误：{exc}")
        pairs, _ = pair_inputs(args.bank.resolve(), args.journal.resolve())
        for index, side in enumerate(("bank", "journal")):
            extra = coverage_evidence.get(side, {}).keys() - {pair[index].name for pair in pairs}
            if extra:
                parser.error(f"{side} 来源凭据未找到对应输入：{'、'.join(sorted(extra))}")
    session = args.session.resolve()
    session.mkdir(parents=True, exist_ok=False)
    result = run_accounts(args.bank.resolve(), args.journal.resolve(), args.output.resolve(), config, session,
                          log=lambda message: print(message, flush=True),
                          bank_columns=bank_columns, journal_columns=journal_columns, matcher_factory=matcher_factory,
                          coverage_evidence=coverage_evidence)
    print("核对结果 " + json.dumps(result, ensure_ascii=False), flush=True)
    return 1 if result["待处理"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
