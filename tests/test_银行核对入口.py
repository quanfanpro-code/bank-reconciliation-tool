from concurrent.futures import ThreadPoolExecutor
from importlib import import_module
import json
import time

import pytest

from llm_assistant import CandidateSemanticRequest, SemanticCandidate


def test_按完整账号配对并排除核对报告与重复输入(tmp_path):
    pair_inputs = import_module('银行核对入口').pair_inputs
    bank, ledger = tmp_path / '流水', tmp_path / '日记账'
    bank.mkdir()
    ledger.mkdir()
    a = bank / '甲公司-甲银行-00000000001.xlsx'
    b = ledger / '甲公司-甲银行-00000000001-日记账.xlsx'
    missing = bank / '甲公司-甲银行-00000000002.xlsx'
    duplicate = bank / '甲公司-甲银行-00000000003.xlsx'
    for p in (a, b, missing, duplicate, duplicate.with_suffix('.csv'),
              ledger / '甲公司-甲银行-00000000003-日记账.xlsx',
              bank / '甲公司-甲银行-00000000001-核对报告.xlsx'):
        p.touch()
    pairs, problems = pair_inputs(bank, ledger)
    assert pairs == [(a, b)]
    assert {p['账号']: p['原因'] for p in problems} == {'00000000002': '缺少日记账', '00000000003': '银行流水有多个文件'}
    assert pair_inputs(a, b) == ([(a, b)], [])


@pytest.mark.parametrize('selected, accepted', [('C-1', True), ('其他任务候选', False)])
def test_只接受本次候选内的判断(tmp_path, selected, accepted):
    reviewer = import_module('银行核对入口').AgentReview(tmp_path, timeout=4)
    request = CandidateSemanticRequest('LLM000001', (SemanticCandidate(
        'C-1', '2026-02-01', '2026-02-02', 100, 100,
        {'摘要': '餐费报销'}, {'摘要': '报销'}, (),
    ), SemanticCandidate('C-2', '2026-02-01', '2026-02-02', 100, 100,
                         {'摘要': '餐费报销'}, {'摘要': '办公费用'}, ())))
    with ThreadPoolExecutor() as pool:
        future = pool.submit(reviewer.evaluate_candidates, request)
        deadline = time.monotonic() + 2
        while not (requests := list(tmp_path.glob('request-*.json'))):
            assert time.monotonic() < deadline
            time.sleep(.02)
        packet = json.loads(requests[0].read_text(encoding='utf-8'))
        assert packet['candidates'][0]['bank_fields'] == {'摘要': '餐费报销'}
        (tmp_path / f"answer-{packet['id']}.json").write_text(json.dumps({
            'id': packet['id'], 'selected_candidate_id': selected,
            'semantic_score': 90, 'reason': '两侧摘要均为报销',
        }, ensure_ascii=False), encoding='utf-8')
        verdict = future.result(timeout=4)
    assert verdict.fallback_used is not accepted
    if accepted:
        assert (verdict.selected_candidate_id, verdict.semantic_score, verdict.reason) == ('C-1', 90, '两侧摘要均为报销')


def test_字段歧义只拦受影响账户且可按明确字段补跑(tmp_path, monkeypatch):
    from decimal import Decimal

    import pandas as pd
    from openpyxl import load_workbook

    from data_structures import LLMDecisionRecord, MatcherConfig

    entry = import_module("银行核对入口")
    monkeypatch.setattr(entry.AgentReview, "evaluate_candidates", lambda _, request: LLMDecisionRecord(
        request.request_id, tuple(candidate.candidate_id for candidate in request.candidates), fallback_used=True,
    ))
    run_accounts = entry.run_accounts
    bank, journal, output = (tmp_path / name for name in ("bank", "journal", "reports"))
    bank.mkdir()
    journal.mkdir()
    dates = ["2026-01-01", "2026-01-02", "2026-01-03"]
    for account in ("000001", "000002"):
        name = f"公司-银行-{account}"
        content = {"交易日期": dates, "交易金额（元）": [100, -40, 50], "账户": [account] * 3}
        if account == "000002":
            content["交易金额(元)"] = [1, 2, 3]
        pd.DataFrame(content).to_excel(bank / (name + ".xlsx"), index=False)
        pd.DataFrame({
            "凭证日期": dates, "借方金额": [100, 0, 50], "贷方金额": [0, 40, 0], "账户": [account] * 3,
        }).to_excel(journal / (name + "-日记账.xlsx"), index=False)
    result = run_accounts(bank, journal, output, MatcherConfig(), tmp_path / "session")
    assert len(result["完成"]) == len(result["待处理"]) == 1
    assert len(list(output.glob("*.xlsx"))) == 1
    blocked = result["待处理"][0]
    assert blocked["字段问题"][0]["候选列"]["amount"] == ["交易金额(元)", "交易金额（元）"]
    retry = run_accounts(
        bank / "公司-银行-000002.xlsx", journal / "公司-银行-000002-日记账.xlsx",
        output, MatcherConfig(), tmp_path / "retry", bank_columns={"amount": "交易金额（元）"},
    )
    assert not retry["待处理"]
    book = load_workbook(retry["完成"][0]["报告"], read_only=True, data_only=True)
    assert book.sheetnames == ["核对总览", "月度核对", "日度核对", "待核对", "待核对明细", "交易明细"]
    amounts = [Decimal(str(row[3])) - Decimal(str(row[4])) for row in list(book["交易明细"].values)[1:]
               if row[1] == "银行"]
    book.close()
    assert amounts == [100, -40, 50]
