"""用边界数据和真实 Excel 验证写出结果，避免读写两端共享错误仍通过。"""
from types import SimpleNamespace
from contextlib import nullcontext
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd
import pytest

from 排版计划 import 簿计划, 表计划, 格样式
from 报告写入 import 渲染簿计划
from tests.test_人工结果计算 import excel

命名空间 = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def test_引号属性和单引号表名可被Excel解析(tmp_path):
    表 = 表计划(名称="甲'乙\"表", 表=pd.DataFrame({'内容': ['原文']}),
               覆盖={(2, 1): 格样式(批注=('说明', '复核"人'), 超链接='#\'甲\'\'乙"表\'!A1')},
               验证=[{'type': 'list', 'sqref': 'A2', 'formula1': '"是,否"', 'error': '请选择"是"或"否"'}],
               打印={'title_rows': '1:1', 'area': 'A1:A2'})
    文件 = tmp_path / '引号.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    with zipfile.ZipFile(文件) as 包:
        for 名 in 包.namelist():
            if 名.endswith(('.xml', '.rels', '.vml')):
                ET.fromstring(包.read(名))
        簿 = ET.fromstring(包.read('xl/workbook.xml'))
        assert 簿.find(命名空间+'sheets')[0].get('name') == 表.名称
        名称们 = 簿.find(命名空间+'definedNames')
        assert all("甲''乙\"表" in n.text for n in 名称们)


def test_连续链接合并且缺口和不同目标保留(tmp_path):
    表 = 表计划(名称='证据', 表=pd.DataFrame({'跳转': ['查看']*5, '其他': ['查看']*5}),
               覆盖={(2, 1): 格样式(超链接='#证据!A1'), (3, 1): 格样式(超链接='#证据!A1'),
                    (5, 1): 格样式(超链接='#证据!A1'), (6, 1): 格样式(超链接='#证据!B1'),
                    (2, 2): 格样式(超链接='#证据!A1'), (3, 2): 格样式(超链接='#证据!A1')})
    文件 = tmp_path / '连续链接.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    with zipfile.ZipFile(文件) as 包:
        根 = ET.fromstring(包.read('xl/worksheets/sheet1.xml'))
        refs = {n.get('ref') for n in 根.find(命名空间+'hyperlinks')}
        assert refs == {'A2:A3', 'B2:B3', 'A5', 'A6'}
    from 直读XML import 读簿
    with nullcontext(读簿(文件)) as 簿:
        assert 簿['证据'].超链接[(3, 1)] == '#证据!A1'
        assert (4, 1) not in 簿['证据'].超链接
        assert 簿['证据'].超链接[(6, 1)] == '#证据!B1'


def test_列上的批注和超链接不被快速通道丢弃(tmp_path, monkeypatch):
    import 排版计划
    from 排版计划 import 列计划
    monkeypatch.setattr(排版计划, '_快速通道行数', 1)
    表 = 表计划(名称='证据', 表=pd.DataFrame({'内容': ['甲', '乙']}),
               列=[列计划(基础样式=格样式(批注=('要核查', '复核'), 超链接='#证据!A1'))])
    文件 = tmp_path / '列附属.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    from 直读XML import 读簿
    with nullcontext(读簿(文件)) as 簿:
        assert 簿['证据'].批注[(3, 1)] == ('要核查', '复核')
        assert 簿['证据'].超链接[(3, 1)] == '#证据!A1'


def test_不可合并的链接超限明确拒绝写出(tmp_path, monkeypatch):
    import 直写XML
    monkeypatch.setattr(直写XML, '_超链接上限', 2, raising=False)
    表 = 表计划(名称='证据', 表=pd.DataFrame({'跳转': ['查看']*3}),
               覆盖={(r, 1): 格样式(超链接=f'#证据!A{r}') for r in range(2, 5)})
    with pytest.raises(ValueError, match='超链接.*上限'):
        渲染簿计划(簿计划(表=[表]), str(tmp_path / '超限.xlsx'))


@pytest.mark.parametrize('值', ['含\x01非法字符', '文'*32768, '='+'1+'*4096+'1'], ids=['控制字符', '超长文本', '超长公式'])
def test_非法或超长内容明确报出单元格(tmp_path, 值):
    表 = 表计划(名称='证据', 表=pd.DataFrame({'内容': [值]}))
    with pytest.raises(ValueError, match='证据.*A2'):
        渲染簿计划(簿计划(表=[表]), str(tmp_path / '非法内容.xlsx'))


def test_正式生成入口在交付前验证Excel且失败不交付(tmp_path, monkeypatch):
    from reporter import Reporter
    import Excel验证
    输出 = tmp_path / '正式报告.xlsx'
    顺序 = []
    def 写入(self, path, *args):
        顺序.append('写入')
        表 = 表计划(名称='原始记录', 表=pd.DataFrame({'金额': [100]}))
        渲染簿计划(簿计划(表=[表]), path)
    def 拒绝(path, **kwargs):
        assert not 输出.exists()
        顺序.append('Excel验证')
        raise RuntimeError('Excel打开失败')
    monkeypatch.setattr(Reporter, '_generate_report', 写入)
    monkeypatch.setattr(Excel验证, '验证Excel报告', 拒绝)
    with pytest.raises(RuntimeError, match='Excel打开失败'):
        Reporter(SimpleNamespace()).generate_report(str(输出))
    assert 顺序 == ['写入', 'Excel验证']
    assert not 输出.exists()


def test_真实Excel可打开超过65530个单元格链接并缓存公式(excel, tmp_path):
    from Excel验证 import 验证Excel报告
    数量 = 65531
    表 = 表计划(名称='大表', 表=pd.DataFrame({'跳转': ['返回']*数量, '金额': [1]*数量}),
               覆盖={(r, 1): 格样式(超链接='#大表!A1') for r in range(2, 数量+2)})
    表.表.loc[0, '跳转'] = '=SUM(B2:B65532)'
    文件 = tmp_path / '超过六万链接.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    结果 = 验证Excel报告(文件)
    assert 结果['工作表数'] == 1
    from 直读XML import 读簿
    with nullcontext(读簿(文件, 数据=True)) as 簿:
        assert 簿['大表'].值(2, 1) == 数量
    book = excel.Workbooks.Open(str(文件), UpdateLinks=0, ReadOnly=True)
    try:
        最后一格 = book.Worksheets('大表').Range('A65532')
        assert 最后一格.Hyperlinks.Count == 1
        assert 最后一格.Hyperlinks.Item(1).SubAddress == '大表!A1'
    finally:
        book.Close(SaveChanges=False)


def test_真实Excel重算错误阻止验收(tmp_path):
    from Excel验证 import 验证Excel报告
    表 = 表计划(名称='错误表', 表=pd.DataFrame({'公式': ['=1/0']}))
    文件 = tmp_path / '计算错误.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    with pytest.raises(RuntimeError, match='错误表.*A2.*DIV/0'):
        验证Excel报告(文件)


def test_验证不干扰已有Excel连接(excel, tmp_path):
    from Excel验证 import 验证Excel报告
    原窗口 = excel.Hwnd
    表 = 表计划(名称='验证', 表=pd.DataFrame({'金额': [100]}))
    文件 = tmp_path / '隔离验证.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    验证Excel报告(文件)
    assert excel.Hwnd == 原窗口
    book = excel.Workbooks.Open(str(文件), ReadOnly=True)
    try:
        assert book.Worksheets('验证').Range('A2').Value == 100
    finally:
        book.Close(SaveChanges=False)


def test_默认样式不被表头覆盖(tmp_path):
    表 = 表计划(名称='验证', 表=pd.DataFrame({'内容': ['甲']}), 表头样式=格样式(加粗=True, 填充色='FF0000'))
    文件 = tmp_path / '默认样式.xlsx'
    渲染簿计划(簿计划(表=[表]), str(文件))
    with zipfile.ZipFile(文件) as 包:
        根 = ET.fromstring(包.read('xl/styles.xml'))
        默认 = 根.find(命名空间+'cellXfs')[0]
        assert 默认.get('fontId') == '0'
        assert 默认.get('fillId') == '0'


def test_验证结束释放自己的Excel进程(tmp_path, monkeypatch):
    import psutil
    import win32com.client
    import win32process
    from Excel验证 import 验证Excel报告
    原启动 = win32com.client.DispatchEx
    创建进程 = []
    def 启动(*args, **kwargs):
        app = 原启动(*args, **kwargs)
        创建进程.append(win32process.GetWindowThreadProcessId(app.Hwnd)[1])
        return app
    monkeypatch.setattr(win32com.client, 'DispatchEx', 启动)
    文件 = tmp_path / '释放验证.xlsx'
    渲染簿计划(簿计划(表=[表计划(名称='验证', 表=pd.DataFrame({'数': [1]}))]), str(文件))
    验证Excel报告(文件)
    assert len(创建进程) == 1
    if psutil.pid_exists(创建进程[0]):
        psutil.Process(创建进程[0]).wait(timeout=10)
