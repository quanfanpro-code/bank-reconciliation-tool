"""由独立 Excel 进程计算人工结果，按原始业务金额和实际对应验收。"""
from copy import deepcopy
from contextlib import contextmanager

import pytest

from data_structures import MatcherConfig, ProcessingStatus, RiskLevel
from matcher import Matcher
from matching_policy import build_group_metrics
from reporter import Reporter
from tests.test_第三阶段特殊业务 import _df


@pytest.fixture(scope='module')
def excel():
    pythoncom = pytest.importorskip('pythoncom')
    client = pytest.importorskip('win32com.client')
    pythoncom.CoInitialize()
    try:
        app = client.DispatchEx('Excel.Application')
    except Exception as exc:
        pythoncom.CoUninitialize()
        pytest.skip(f'本机 Excel 不可用：{exc}')
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.IgnoreRemoteRequests = True
        yield app
    finally:
        app.Quit()
        pythoncom.CoUninitialize()


def _报告(tmp_path, bank_amounts, journal_amounts, selected, alternatives=()):
    def frame(amounts):
        return _df([('2026-01-10', value, {'摘要':'设备款','对方':'甲公司','业务编号':f'P{i}'}) for i,value in enumerate(amounts)])
    config = MatcherConfig(performance_materiality=50, clearly_trivial_threshold=1)
    m = Matcher(frame(bank_amounts), frame(journal_amounts), config, logger=lambda _:None)
    m.run()
    template = m.selected_candidates[0]
    def candidate(identity, banks, journals):
        c = deepcopy(template)
        c.candidate_id, c.final_match_id = 'C'+identity, identity
        c.bank_idxs, c.journal_idxs = tuple(banks), tuple(journals)
        c.bank_dates = tuple(m.bank.loc[list(banks),'date'])
        c.journal_dates = tuple(m.journal.loc[list(journals),'date'])
        c.metrics = build_group_metrics(list(m.bank.loc[list(banks),'amount_decimal']),list(m.journal.loc[list(journals),'amount_decimal']))
        c.processing_status, c.risk_level = ProcessingStatus.FLAGGED, RiskLevel.HIGH
        c.is_ambiguous = True
        c.processing_reason = '已有候选需核对原始业务依据'
        c.evidence = {}
        return c
    m.selected_candidates = [candidate(f'M{i+1}',banks,journals) for i,(banks,journals) in enumerate(selected)]
    m.candidates = list(m.selected_candidates)
    m.difference_pools = []
    m.business_events = []
    for owner,banks,journals in alternatives:
        c = candidate(f'A{len(m.candidates)}',banks,journals)
        m.candidates.append(c)
        m.selected_candidates[owner].evidence.setdefault('alternative_candidate_ids',[]).append(c.candidate_id)
    for source, attr in [('bank','bank_idxs'),('journal','journal_idxs')]:
        table = getattr(m,source)
        table['matched'], table['match_id'] = False, ''
        for c in m.selected_candidates:
            indexes = list(getattr(c,attr))
            table.loc[indexes,'matched'] = True
            table.loc[indexes,'match_id'] = c.final_match_id
            table.loc[indexes,'processing_status'] = c.processing_status.value
            table.loc[indexes,'risk_level'] = c.risk_level.value
    path = tmp_path/'人工结果.xlsx'
    Reporter(m).generate_report(str(path),config=config)
    return path


@contextmanager
def _打开(excel, path):
    book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
    try:
        excel.CalculateFullRebuild()
        yield book
    finally:
        book.Close(SaveChanges=False)


def _表(book,name):
    values = book.Worksheets(name).UsedRange.Value
    return [dict(zip(values[0], row)) for row in values[1:]]


def _结论(book):
    return {r['项目']:r['数值'] for r in _表(book,'核对结论')}


def _选择(excel,book,identity,value):
    for name in ['待处理事项']:
        sheet=book.Worksheets(name)
        values=sheet.UsedRange.Value
        headers=list(values[0])
        for number,row in enumerate(values[1:],2):
            data=dict(zip(headers,row))
            if data.get('事项编号')==identity:
                sheet.Cells(number,headers.index('核查结论')+1).Value=value
                excel.CalculateFullRebuild()
                return
    raise AssertionError(f'找不到人工填写事项：{identity}')


def _明细(book):
    return {row['原始记录号']:row for row in _表(book,'核对明细')}


def _覆盖(book,bank,journal):
    summary=_结论(book)
    assert summary['银行当前已对平金额覆盖率']==pytest.approx(bank)
    assert summary['序时账当前已对平金额覆盖率']==pytest.approx(journal)


def test_未填及自由结论均保留真实差额与程序配对(excel,tmp_path):
    path=_报告(tmp_path,[100],[90],[((0,),(0,))])
    with _打开(excel,path) as book:
        summary=_结论(book)
        assert summary['人工全查未核对数']==1
        assert summary['待处理事项数']==1
        assert summary['已处理事项数']==0
        assert '待人工' in _表(book,'复核事项索引')[0]['当前核对结论']
        original={key:value for key,value in _表(book,'月度核对')[0].items() if key in ['银行-收入金额','序时账-收入金额','收入金额差额','支出金额差额']}
        assert original['收入金额差额']==10
        _选择(excel,book,'M1','已核查，差额为手续费，凭证另附')
        assert _结论(book)['待处理事项数']==0
        assert _结论(book)['已处理事项数']==1
        _覆盖(book,0,0)
        monthly=_表(book,'月度核对')[0]
        assert all(monthly[key]==value for key,value in original.items())
        assert monthly['已确认收入差']==0
        assert monthly['未确认收入差']==10
        assert all(not row['当前对应'] for row in _明细(book).values())
        _选择(excel,book,'M1','发现差异，尚待补充资料')
        assert all(not row['当前对应'] and '已记录核查结论' in row['当前核对结果'] for row in _明细(book).values())
        _覆盖(book,0,0)
        assert _表(book,'月度核对')[0]['未确认收入差']==10


def test_结论提及其他凭证也不自动采用或释放原始行(excel,tmp_path):
    path=_报告(tmp_path,[100],[100,100],[((0,),(0,))],[(0,(0,),(1,))])
    with _打开(excel,path) as book:
        untouched=_明细(book)
        assert len(untouched)==3 and all(not row['当前对应'] for row in untouched.values())
        single_id=untouched['账面3']['事项编号']
        single_status=next(row['当前核对结论'] for row in _表(book,'复核事项索引') if row['事项编号']==single_id)
        month_statuses=[row['当前核对结论'] for row in _表(book,'月度差异组成')]
        initial_state={
            '未对应原始记录数':_结论(book)['当前未对应原始记录数'],
            '单边待人工':'待人工' in single_status,
            '月度仍有未对应':bool(month_statuses) and all('仍有未对应' in value for value in month_statuses),
        }
        assert initial_state=={'未对应原始记录数':3,'单边待人工':True,'月度仍有未对应':True}, (initial_state,single_status,month_statuses)
        _选择(excel,book,'M1','原凭证有误，实际应查另一张凭证')
        rows=_明细(book)
        assert all(not row['当前对应'] for row in rows.values())
        assert '已记录核查结论' in rows['账面2']['当前核对结果']
        assert '已确认' not in rows['账面2']['当前核对结果']
        _覆盖(book,0,0)
        contributions=_表(book,'月度差异组成')
        assert len(contributions)==2
        initial=next(row for row in contributions if row['事项编号']=='M1')
        assert initial['收入差额']==0
        assert initial['已确认收入差']==0
        assert initial['未确认收入差']==0
        single=rows['账面3']['事项编号']
        _选择(excel,book,single,'已查明，重复记账待处理')
        assert not _明细(book)['账面3']['当前对应']
        item=next(row for row in _表(book,'复核事项索引') if row['事项编号']==single)
        assert '已记录核查结论' in item['当前核对结论']
        _覆盖(book,0,0)


def test_两项结论不会导致交易重新分配或重复占用(excel,tmp_path):
    path=_报告(tmp_path,[100,200],[100,200],[((0,),(0,)),((1,),(1,))],[(1,(0,),(1,))])
    with _打开(excel,path) as book:
        _选择(excel,book,'M1','已查原始凭证，详见备注')
        _选择(excel,book,'M2','已查另一张凭证，详见备注')
        assert _结论(book)['已处理事项数']==2
        assert all(not row['当前对应'] for row in _明细(book).values())
        assert all('已记录核查结论' in row['当前核对结论'] for row in _表(book,'复核事项索引'))
        _覆盖(book,0,0)


def test_人工表整行排序后填写不串事项且完整组成可追溯(excel,tmp_path):
    path=_报告(tmp_path,[40,60,200],[100,200],[((0,1),(0,)),((2,),(1,))])
    with _打开(excel,path) as book:
        sheet=book.Worksheets('待处理事项')
        rows=_表(book,'待处理事项')
        assert all(row['事项编号'] for row in rows)
        before=[row['事项编号'] for row in rows]
        headers=list(sheet.UsedRange.Value[0])
        # 只排序数据区，避免Excel把中文表头作为普通记录参与排序。
        sheet.Range(sheet.Cells(2, 1), sheet.Cells(sheet.UsedRange.Rows.Count, len(headers))).Sort(
            Key1=sheet.Cells(2,headers.index('事项编号')+1), Order1=1, Header=2)
        after=[row['事项编号'] for row in _表(book,'待处理事项')]
        assert after!=before
        assert after==sorted(before)
        _选择(excel,book,'M1','组内记录已查，见备注')
        _选择(excel,book,'M2','尚有差异')
        details=_明细(book)
        assert len(details)==5 and all(not row['当前对应'] for row in details.values())
        index={row['事项编号']:row for row in _表(book,'复核事项索引')}
        assert index['M1']['人工核对结果']=='组内记录已查，见备注'
        assert index['M2']['人工核对结果']=='尚有差异'
        _覆盖(book,0,0)


def test_粘贴自由文字可记录但不增加覆盖(excel,tmp_path):
    path=_报告(tmp_path,[100],[100],[((0,),(0,))])
    with _打开(excel,path) as book:
        _选择(excel,book,'M1','已向企业询问，等待回单')
        assert '已记录核查结论' in _表(book,'复核事项索引')[0]['当前核对结论']
        assert _结论(book)['人工全查未核对数']==0
        assert _结论(book)['待处理事项数']==0
        assert _结论(book)['已处理事项数']==1
        assert all(not row['当前对应'] for row in _明细(book).values())
        _覆盖(book,0,0)
