"""疑点供查阅，人工结论是自由文字，不能重组交易或提高覆盖率。"""
from copy import deepcopy
from types import SimpleNamespace

import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus, RiskLevel
from matcher import Matcher
from matching_policy import build_group_metrics
from reporter import Reporter
from 排版计划 import 生成簿计划, 簿计划, 表计划
from 报告写入 import 渲染簿计划
from tests.test_第三阶段特殊业务 import _df


def _业务数据():
    def frame(amounts):
        return _df([('2026-01-10', a, {'摘要': '设备款', '业务编号': f'P{i}'})
                    for i, a in enumerate(amounts)])
    config = MatcherConfig(performance_materiality=50, clearly_trivial_threshold=1)
    m = Matcher(frame([100, 200, 300]), frame([100, 200, 300, 100]), config, logger=lambda _: None)
    m.run()
    template = m.selected_candidates[0]
    def candidate(identity, b, j, auto=False):
        c = deepcopy(template)
        c.candidate_id, c.final_match_id = 'C'+identity, identity
        c.bank_idxs, c.journal_idxs = (b,), (j,)
        c.bank_dates, c.journal_dates = (m.bank.loc[b, 'date'],), (m.journal.loc[j, 'date'],)
        c.metrics = build_group_metrics(m.bank.loc[[b], 'amount_decimal'].tolist(), m.journal.loc[[j], 'amount_decimal'].tolist())
        c.processing_status = ProcessingStatus.AUTO_CONFIRMED if auto else ProcessingStatus.FLAGGED
        c.risk_level = RiskLevel.NORMAL if auto else RiskLevel.HIGH
        c.is_ambiguous = not auto
        c.processing_reason = '程序建议，须查原始凭证'
        c.evidence = {}
        return c
    m.selected_candidates = [candidate('M1', 0, 0), candidate('M2', 1, 1), candidate('M3', 2, 2, True)]
    alternative = candidate('A1', 0, 3)
    m.selected_candidates[0].evidence['alternative_candidate_ids'] = [alternative.candidate_id]
    m.candidates = [*m.selected_candidates, alternative]
    m.difference_pools, m.business_events = [], []
    for side, field in [('bank', 'bank_idxs'), ('journal', 'journal_idxs')]:
        table = getattr(m, side)
        table['matched'], table['match_id'] = False, ''
        for c in m.selected_candidates:
            ids = list(getattr(c, field))
            table.loc[ids, 'matched'] = True
            table.loc[ids, 'match_id'] = c.final_match_id
            table.loc[ids, 'processing_status'] = c.processing_status.value
            table.loc[ids, 'risk_level'] = c.risk_level.value
    reporter = Reporter(m)
    return reporter, config, reporter.build_report_tables(config)


def test_人工区自由填写并保留黄色入口():
    reporter, config, tables = _业务数据()
    plan = 生成簿计划(reporter, tables, config)
    pending = plan.取表('待处理事项')
    assert '核查结论' in pending.表
    assert not pending.验证
    assert pending.表['核查结论'].isna().all()
    assert pending.表['备注'].isna().all()
    for row in range(2, len(pending.表)+2):
        assert pending.后覆盖[(row, 8)].填充色 == 'FFF2CC'
        assert pending.后覆盖[(row, 10)].填充色 == 'FFF2CC'
        assert pending.覆盖[(row, 11)].超链接.startswith("#'事项证据'!")


def test_候选仅作线索不进入可人工生效选项():
    _, _, tables = _业务数据()
    assert len(tables['复核候选选项']) == 3
    clues = tables['其他对应线索']
    assert set(clues['线索编号']) == {'参考线索1'}
    assert set(clues['来源']) == {'银行流水', '序时账'}
    assert clues['原始出处'].str.contains('行').all()
    questions = '\n'.join(tables['复核事项索引']['具体问题'])
    assert '选对应方案' not in questions and '选择保留未对应' not in questions


def test_参考线索可直接查看而非跳到隐藏表():
    reporter, config, tables = _业务数据()
    plan = 生成簿计划(reporter, tables, config)
    assert plan.取表('其他对应线索').状态 == 'visible'


def test_交付拒绝重新引入配对下拉(tmp_path):
    from Excel验证 import 检查人工填写区
    path = tmp_path/'错误配对菜单.xlsx'
    sheet = 表计划(名称='待处理事项', 表=pd.DataFrame({'核查结论': [''], '当前结果': ['待核对'], '备注': ['']}))
    sheet.验证.append({'type': 'list', 'formula1': '"确认对应,采用其他对应1"', 'sqref': 'A2'})
    渲染簿计划(簿计划(表=[sheet]), str(path))
    with pytest.raises(RuntimeError, match='填写区.*下拉'):
        检查人工填写区(path)


def test_Excel自由结论更新进度但不改配对和金额(excel, tmp_path):
    from tests.test_人工结果计算 import _打开, _表, _结论
    reporter, config, _ = _业务数据()
    path = tmp_path/'自由核查.xlsx'
    reporter.generate_report(str(path), config=config)
    with _打开(excel, path) as book:
        sheet = book.Worksheets('待处理事项')
        rows = _表(book, '待处理事项')
        headers = list(sheet.UsedRange.Value[0])
        inputs = {r['事项编号']: i for i, r in enumerate(rows, 2)}
        before = _表(book, '核对明细')
        monthly = _表(book, '月度核对')
        summary = _结论(book)
        count = len(rows)
        for identity, text in [('M1', '已核查，真正对应在其他凭证中，见备注'), ('M2', '仍有差异，待补资料')]:
            row = inputs[identity]
            sheet.Cells(row, headers.index('核查结论')+1).Value = text
            sheet.Cells(row, headers.index('备注')+1).Value = '凭证号不在程序候选清单中'
        excel.CalculateFullRebuild()
        after = _表(book, '核对明细')
        fields = ('原始记录号', '金额', '当前对应', '已对平金额', '已确认收入差', '已确认支出差')
        assert [tuple(r[k] for k in fields) for r in after] == [tuple(r[k] for k in fields) for r in before]
        assert _表(book, '月度核对') == monthly
        for name in ('银行当前已对平金额覆盖率', '序时账当前已对平金额覆盖率', '当前未对应原始记录数'):
            assert _结论(book)[name] == summary[name]
        assert _结论(book)['已处理事项数'] == 2
        assert _结论(book)['待处理事项数'] == count-2
        assert all('已记录核查结论' in r['当前结果'] for r in _表(book, '待处理事项') if r['事项编号'] in ('M1', 'M2'))
        sheet.Cells(inputs['M1'], headers.index('核查结论')+1).ClearContents()
        excel.CalculateFullRebuild()
        assert _结论(book)['已处理事项数'] == 1
        assert _结论(book)['待处理事项数'] == count-1


from tests.test_人工结果计算 import excel
