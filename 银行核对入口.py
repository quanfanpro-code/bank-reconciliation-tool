"""为 Agent 提供单户及目录核对入口。用户设置由 Agent 在问答中收集。"""

import argparse
from collections import defaultdict
from dataclasses import asdict, replace
from datetime import datetime
from decimal import Decimal
import json
from pathlib import Path
import re
import time
from uuid import uuid4

from data_structures import MatcherConfig
from 列识别 import ColumnRecognitionError, FIELD_LABELS


def _identity(path):
    stem = re.sub(r"[-_](?:银行存款日记账|银行日记账|银行存款序时账|日记账|序时账)$", "", path.stem)
    parts = stem.rsplit("-", 2)
    if len(parts) == 3 and all(part.strip() for part in parts) and parts[-1].isascii() and parts[-1].isdigit():
        return tuple(part.strip() for part in parts)
    return None


def pair_inputs(bank, journal):
    if bank.is_file() and journal.is_file():
        if bank.resolve() == journal.resolve():
            raise ValueError("银行流水和日记账不能使用同一个文件")
        left, right = _identity(bank), _identity(journal)
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
            identity = _identity(path)
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


class AgentReview:
    """在本次任务目录交接候选和 Agent 答复。"""

    def __init__(self, directory, timeout=120, sources=()):
        self.directory = directory
        self.timeout = timeout
        self.sources = sources
        self.timed_out = False
        directory.mkdir(parents=True, exist_ok=True)

    def evaluate_candidates(self, request):
        from data_structures import LLMDecisionRecord
        from llm_assistant import ONLINE_ALLOWED_FIELDS

        ids = tuple(candidate.candidate_id for candidate in request.candidates)
        verdict = LLMDecisionRecord(request.request_id, ids, provider="当前 Agent", fallback_used=True)
        if len(ids) < 2:
            return verdict
        if self.timed_out:
            verdict.error = "当前 Agent 未在等待时间内返回答复"
            return verdict

        token = uuid4().hex
        candidates = []
        for candidate in request.candidates:
            summary = asdict(candidate)
            for key in ("bank_fields", "journal_fields"):
                summary[key] = {name: value for name, value in summary[key].items() if name in ONLINE_ALLOWED_FIELDS}
            candidates.append(summary)
        request_path = self.directory / f"request-{token}.json"
        staged = request_path.with_suffix(".new")
        staged.write_text(json.dumps({"id": token, "sources": self.sources, "candidates": candidates}, ensure_ascii=False), encoding="utf-8")
        staged.replace(request_path)
        response_path = self.directory / f"answer-{token}.json"
        deadline = time.monotonic() + self.timeout
        while not response_path.is_file() and time.monotonic() < deadline:
            time.sleep(.1)
        if not response_path.is_file():
            self.timed_out = True
            verdict.error = "当前 Agent 未在等待时间内返回答复"
            return verdict

        try:
            answer = json.loads(response_path.read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError, ValueError):
            answer = None
        if not isinstance(answer, dict) or answer.get("id") != token:
            verdict.error = "Agent 答复无法关联到本次候选"
            return verdict
        selected = answer.get("selected_candidate_id")
        score = answer.get("semantic_score")
        reason = answer.get("reason")
        if selected in ids and type(score) is int and 0 <= score <= 100 and isinstance(reason, str) and reason.strip():
            return replace(verdict, selected_candidate_id=selected, semantic_score=score,
                           reason=reason.strip(), fallback_used=False)
        verdict.error = "当前 Agent 保留疑点" if not selected else "Agent 返回了无效的候选或依据"
        return verdict


def run_accounts(bank, journal, output, config, session, *, log=print,
                 bank_columns=None, journal_columns=None):
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
    for bank_file, journal_file in pairs:
        try:
            target = output / f"{bank_file.stem}_核对报告_{suffix}.xlsx"

            def set_reviewer(matcher):
                matcher.llm_assistant = AgentReview(session / "复核", sources=(str(bank_file), str(journal_file)))

            report = run_reconciliation(
                str(bank_file), str(journal_file),
                {"_column_auto": True, "_column_overrides": bank_columns},
                {"_column_auto": True, "_column_overrides": journal_columns},
                config, llm_config=LLMConfig(enabled=False), logger=log,
                output_path=target, project_store=store, matcher_ready=set_reviewer,
            )
            results.append({"银行流水": str(bank_file), "日记账": str(journal_file), "报告": str(report)})
        except ColumnRecognitionError as exc:
            problems.append({
                "银行流水": str(bank_file), "日记账": str(journal_file),
                "原因": str(exc), "字段问题": exc.issues,
            })
        except Exception as exc:
            problems.append({"银行流水": str(bank_file), "日记账": str(journal_file), "原因": str(exc)})
    return {"完成": results, "待处理": problems}


def _column_overrides(parser, entries):
    result = {}
    for entry in entries:
        role, separator, column = entry.partition("=")
        if not separator or role not in FIELD_LABELS or not column or role in result:
            parser.error(f"无效或重复的字段设置：{entry}")
        result[role] = column
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bank", required=True, type=Path)
    parser.add_argument("--journal", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--session", required=True, type=Path)
    parser.add_argument("--trivial", type=Decimal)
    parser.add_argument("--materiality", type=Decimal)
    parser.add_argument("--bank-column", action="append", default=[], metavar="字段=原列名")
    parser.add_argument("--journal-column", action="append", default=[], metavar="字段=原列名")
    args = parser.parse_args()
    bank_columns = _column_overrides(parser, args.bank_column)
    journal_columns = _column_overrides(parser, args.journal_column)
    config = replace(MatcherConfig(), **{
        key: value for key, value in {"clearly_trivial_threshold": args.trivial, "performance_materiality": args.materiality}.items()
        if value is not None
    })
    session = args.session.resolve()
    session.mkdir(parents=True, exist_ok=False)
    result = run_accounts(args.bank.resolve(), args.journal.resolve(), args.output.resolve(), config, session,
                          log=lambda message: print(message, flush=True),
                          bank_columns=bank_columns, journal_columns=journal_columns)
    print("核对结果 " + json.dumps(result, ensure_ascii=False), flush=True)
    return 1 if result["待处理"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
