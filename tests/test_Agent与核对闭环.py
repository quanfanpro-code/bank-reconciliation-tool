from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal
import json
from pathlib import Path
import time

from openpyxl import load_workbook
import pandas as pd
import pytest

from application import run_reconciliation
from data_structures import MatcherConfig
from llm_assistant import CandidateSemanticRequest, SemanticCandidate
from 银行核对入口 import AgentReview, run_accounts


CONFIG = MatcherConfig(clearly_trivial_threshold=20, performance_materiality=500)


def _sources(folder, bank, journal, names=("流水.xlsx", "日记账.xlsx")):
    bank_path, journal_path = (folder / name for name in names)
    pd.DataFrame(bank).to_excel(bank_path, index=False)
    pd.DataFrame(journal).to_excel(journal_path, index=False)
    return bank_path, journal_path


def _table(path, sheet):
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(book[sheet].values)
    finally:
        book.close()


@pytest.mark.parametrize("score,confirmed", [(100, True), (20, False)])
def test_Agent原表依据改变最终结果并写入报告(tmp_path, score, confirmed):
    bank, journal = _sources(tmp_path, {
        "交易日期": ["2026-02-05"], "交易金额": [-1000], "摘要": ["转账"],
        "附言": ["2026年1月天然气款"],
    }, {
        "凭证日期": ["2026-02-05"] * 2, "发生额": [-1000, -1000],
        "摘要": ["2026年1月天然气款", "2025年12月天然气款"],
    })
    session = tmp_path / "任务"
    with ThreadPoolExecutor() as pool:
        job = pool.submit(run_accounts, bank, journal, tmp_path / "报告", CONFIG, session)
        deadline = time.monotonic() + 35
        answered = set()
        while not job.done():
            for request in (session / "复核").glob("request-*.json"):
                if request in answered:
                    continue
                packet = json.loads(request.read_text(encoding="utf-8"))
                selected = next(item for item in packet["candidates"]
                                if item["journal_fields"].get("摘要") == "2026年1月天然气款")
                staged = request.with_suffix(".new")
                staged.write_text(json.dumps({"id": packet["id"], "selected_candidate_id": selected["candidate_id"],
                                              "semantic_score": score,
                                              "reason": "银行附言与账上摘要同为2026年1月天然气款"},
                                             ensure_ascii=False), encoding="utf-8")
                staged.replace(request.with_name(f"answer-{packet['id']}.json"))
                answered.add(request)
            assert time.monotonic() < deadline
            time.sleep(.02)
        result = job.result()
    assert len(answered) == 1 and not result["待处理"]
    workbook = Path(result["完成"][0]["报告"])
    records = _table(workbook, "交易明细")[1:]
    debit = next(row for row in records if row[1] == "银行")
    chosen = next(row for row in records if row[0] == debit[0] and row[1] == "账上")
    if confirmed:
        assert chosen[7] == "2026年1月天然气款"
        assert "Agent 判断依据" in debit[10]
    pending = _table(workbook, "待核对")[1:]
    assert all(row[0] != debit[0] for row in pending) is confirmed
    assert result["Agent复核"] == {"已完成": 1, "超时": 0, "未复核": 0}


def test_待核对显示字段冲突和月度入选原因(tmp_path):
    date = "2026-02-05"
    bank, journal = _sources(tmp_path, {
        "交易日期": [date] * 3, "交易金额": [-1000, -300, -300],
        "摘要": ["设备采购", "第一笔零星付款", "第二笔零星付款"],
        "对方户名": ["甲公司", "丙公司", "丁公司"],
    }, {
        "凭证日期": [date], "发生额": [-1000], "摘要": ["设备采购"], "对方户名": ["乙公司"],
    })
    workbook = run_reconciliation(str(bank), str(journal), None, None, CONFIG,
                                  output_path=tmp_path / "核对.xlsx")
    actions = [row[2] for row in _table(workbook, "待核对")[1:]]
    assert any("对方户名" in message and "冲突" in message for message in actions)
    assert any("月度累计" in message and "实际执行重要性水平" in message for message in actions)
    assert all("综合可信度" not in message for message in actions)


@pytest.mark.parametrize("use_filename", [False, True])
def test_余额配对按表内账号或受支持的账单后缀(tmp_path, use_filename):
    account = "001234567890"
    bank_data = {"交易日期": ["2026-02-03", "2026-02-05", "2026-02-05"], "交易金额": [70, -120, -180],
                 "交易后余额": [1000, 880, 700], "摘要": ["同笔收款", "转账", "转账"]}
    journal_data = {"凭证日期": ["2026-02-03", "2026-02-06"], "发生额": [70, -300],
                    "账面余额": [1000, 700], "摘要": ["同笔收款", "付款"]}
    names = ("流水.xlsx", "日记账.xlsx")
    if use_filename:
        names = (f"甲公司-甲银行-{account}.xlsx", f"甲公司-甲银行-{account}-银行存款日记账.xlsx")
        journal_data["辅助核算名称"] = [f"甲银行-{account}"] * 2
        journal_data["凭证日期"][-1] = "2026-02-05"
    else:
        bank_data["账户"] = [account] * 3
        journal_data["账户"] = [account] * 2
    bank, journal = _sources(tmp_path, bank_data, journal_data, names)
    matchers = []
    run_reconciliation(str(bank), str(journal), None, None, CONFIG,
                       output_path=tmp_path / "核对.xlsx", matcher_ready=matchers.append)
    assert matchers[0].balance_scope_verified
    assert [(item["bank_idxs"], item["journal_idxs"], item["processing_status"])
            for item in matchers[0].matches] == [([0], [0], "自动确认"), ([1, 2], [1], "自动确认")]


def test_日记账收入支出分列保留实际方向(tmp_path):
    dates = ["2026-02-05", "2026-02-06"]
    bank, journal = _sources(tmp_path, {"交易日期": dates, "交易金额": [1000, -80]},
                            {"凭证日期": dates, "收入金额": [1000, None], "支出金额": [None, 80]})
    workbook = run_reconciliation(str(bank), str(journal), None, None, CONFIG,
                                  output_path=tmp_path / "核对.xlsx")
    assert [(Decimal(str(row[3])), Decimal(str(row[4])))
            for row in _table(workbook, "交易明细")[1:] if row[1] == "账上"] == [(1000, 0), (0, 80)]


def test_首次超时后仍发出第二次请求并统计(tmp_path):
    candidate = lambda name: SemanticCandidate(name, "2026-02-01", "2026-02-01", 10, 10, {}, {}, ())
    reviewer = AgentReview(tmp_path, timeout=.12)
    choices = (candidate("甲"), candidate("乙"))
    first = reviewer.evaluate_candidates(CandidateSemanticRequest("first", choices))
    previous = set(tmp_path.glob("request-*.json"))
    assert first.fallback_used
    assert len(previous) == 1
    reviewer.timeout = 3
    with ThreadPoolExecutor() as pool:
        second = pool.submit(reviewer.evaluate_candidates, CandidateSemanticRequest("second", choices))
        deadline = time.monotonic() + 2
        while not (new := set(tmp_path.glob("request-*.json")) - previous) and not second.done():
            assert time.monotonic() < deadline
            time.sleep(.02)
        assert len(new) == 1
        request = new.pop()
        packet = json.loads(request.read_text(encoding="utf-8"))
        staged = request.with_suffix(".new")
        staged.write_text(json.dumps({"id": packet["id"], "selected_candidate_id": "乙",
                                      "semantic_score": 100, "reason": "共同附言"}), encoding="utf-8")
        staged.replace(request.with_name(f"answer-{packet['id']}.json"))
        assert second.result(timeout=4).selected_candidate_id == "乙"
    assert reviewer.review_counts == {"已完成": 1, "超时": 1, "未复核": 0}


@pytest.mark.parametrize("skipped_reason", ["明确冲突", "候选无法区分", "未展示全部候选"])
def test_Agent满分不能确认未排除的其他可能(skipped_reason):
    from data_loader import DataLoader
    from data_structures import LLMDecisionRecord
    from matcher import Matcher

    def read(rows, source):
        frame = pd.DataFrame(rows)
        return DataLoader().standardize_data(frame, {
            "mode": "signed_amount", "date": "日期", "amount": "金额",
            "summary": "摘要", "auxiliary_text_columns": ["摘要", "对方户名"],
        }, source)

    day = "2026-02-05"
    bank = read([{"日期": day, "金额": 1000, "摘要": "转账", "对方户名": "甲公司"}], "bank")
    journals = [{"日期": day, "金额": 1000, "摘要": "一月设备款", "对方户名": ""}]
    journals.append({**journals[0], "摘要": "二月设备款"})
    if skipped_reason == "明确冲突":
        journals[0]["对方户名"] = "乙公司"
    elif skipped_reason == "候选无法区分":
        journals[1]["摘要"] = journals[0]["摘要"]
    else:
        journals.append({**journals[0], "摘要": "三月设备款"})
    journal = read(journals, "journal")

    class Reviewer:
        candidate_limit = 2

        def evaluate_candidates(self, request):
            target = next((candidate for candidate in request.candidates
                           if candidate.journal_fields.get("摘要") == "一月设备款"), request.candidates[0])
            return LLMDecisionRecord(request.request_id, selected_candidate_id=target.candidate_id,
                                     semantic_score=100, reason="Agent 给出判断", provider="当前 Agent")

    matcher = Matcher(bank, journal, CONFIG, llm_assistant=Reviewer())
    matcher.run()
    assert matcher.llm_records and not matcher.llm_records[0].fallback_used
    assert not any(candidate.evidence.get("agent_confirmed") for candidate in matcher.selected_candidates)
