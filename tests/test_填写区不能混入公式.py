"""填写区必须保留业务输入，不能接收程序回显公式或修改缓存的业务表。"""
from types import SimpleNamespace
import pandas as pd
import pytest

from data_structures import MatcherConfig, ProcessingStatus, RiskLevel
from matcher import Matcher
from reporter import Reporter
from tests.test_第三阶段特殊业务 import _df
from 排版计划 import 新建簿计划, 生成簿计划, 簿计划, 表计划
from 报告写入 import 渲染簿计划


def _业务表():
    df = _df([('2026-01-10', 100, {'摘要': '设备款', '业务编号': 'P1'})])
    config = MatcherConfig(performance_materiality=50, clearly_trivial_threshold=1)
    matcher = Matcher(df, df, config, logger=lambda _: None)
    matcher.run()
    for c in matcher.selected_candidates:
        c.processing_status = ProcessingStatus.FLAGGED
        c.risk_level = RiskLevel.HIGH
        c.is_ambiguous = True
    return Reporter(matcher).build_report_tables(config)


def test_排版不修改调用者的业务表且重复生成仍为空白():
    tables = _业务表()
    before = {n: d.copy(deep=True) for n, d in tables.items()}
    生成簿计划(SimpleNamespace(), tables)
    for name in tables:
        pd.testing.assert_frame_equal(tables[name], before[name], obj=name)
    again = 生成簿计划(SimpleNamespace(), tables)
    assert again.取表('待处理事项').表['核查结论'].isna().all()


@pytest.mark.parametrize('前缀', ['', "'"])
def test_旧回显公式不得迁入人工输入(前缀):
    tables = _业务表()
    mirror = 前缀 + '=IF(IFERROR(INDEX(\'待处理事项\'!$H$2:$H$9,MATCH(S2,\'待处理事项\'!$L$2:$L$9,0)),"")="","",1)'
    mask = tables['人工全查']['行别'] == '核对事项'
    tables['人工全查'].loc[mask, '人工核对结果'] = mirror
    tables['人工全查'].loc[mask, '备注'] = mirror.replace('$H$', '$J$')
    pending = 生成簿计划(SimpleNamespace(), tables).取表('待处理事项').表
    assert pending['核查结论'].isna().all()
    assert pending['备注'].isna().all()


def test_真实人工选择与备注保留():
    tables = _业务表()
    mask = tables['人工全查']['行别'] == '核对事项'
    tables['人工全查'].loc[mask, '人工核对结果'] = '确认对应'
    tables['人工全查'].loc[mask, '备注'] = '已核对原始凭证'
    pending = 生成簿计划(SimpleNamespace(), tables).取表('待处理事项').表
    assert pending.iloc[0]['核查结论'] == '确认对应'
    assert pending.iloc[0]['备注'] == '已核对原始凭证'


@pytest.mark.parametrize('错误值', ['=1+1', "'=IF(1=1,1,0)"])
def test_交付入口拒绝填写区公式及公式文本(tmp_path, monkeypatch, 错误值):
    import win32com.client
    from Excel验证 import 验证Excel报告
    def 不应启动(*args, **kwargs):
        raise AssertionError('不应调用Excel')
    monkeypatch.setattr(win32com.client, 'DispatchEx', 不应启动)
    path = tmp_path / '错误填写区.xlsx'
    表 = 表计划(名称='待处理事项', 表=pd.DataFrame({'人工核对结果': [错误值], '当前结果': ['选择无效'], '备注': ['']}))
    渲染簿计划(簿计划(表=[表]), str(path))
    with pytest.raises(RuntimeError, match='填写区.*包含公式'):
        验证Excel报告(path)


@pytest.mark.parametrize('字段,值', [('备注', "'=IF(INDEX('待处理事项'!J2,1)=\"\",\"\",1)"), ('当前结果', '选择无效，请重选')])
def test_交付拒绝回显备注和无效业务状态(tmp_path, 字段, 值):
    from Excel验证 import 检查人工填写区
    path = tmp_path / '业务状态.xlsx'
    数据 = {'人工核对结果': [''], '当前结果': ['待核对'], '备注': ['']}
    数据[字段] = [值]
    渲染簿计划(簿计划(表=[表计划(名称='待处理事项', 表=pd.DataFrame(数据))]), str(path))
    with pytest.raises(RuntimeError, match='填写区'):
        检查人工填写区(path)
