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
