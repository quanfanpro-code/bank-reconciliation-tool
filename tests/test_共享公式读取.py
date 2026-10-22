"""真实 Excel 会把连续公式保存为共享公式，读取和再次写出不能丢掉后续公式。"""
from contextlib import nullcontext
import re
from zipfile import ZipFile, ZIP_DEFLATED

import pandas as pd
import pytest

from 排版计划 import 簿计划, 表计划
from 报告写入 import 渲染簿计划
from 直读XML import 读簿


@pytest.mark.parametrize('公式,起,止,预期', [
    ('A1+$A1+A$1+$A$1', 'B2', 'C4', 'B3+$A3+B$1+$A$1'),
    ('SUM(A:C,$A:$C,1:3,$1:$3)', 'B2', 'C4', 'SUM(B:D,$A:$C,3:5,$1:$3)'),
    ('IF(A1="A1",LOG10(B2),\'甲\'\'乙A1\'!C3)', 'A1', 'A2', 'IF(A2="A1",LOG10(B3),\'甲\'\'乙A1\'!C4)'),
    ('A1', 'B1', 'A1', '#REF!'),
])
def test_共享公式平移边界(公式, 起, 止, 预期):
    from 直读XML import _平移共享公式
    assert _平移共享公式(公式, 起, 止) == 预期


def test_共享公式保持相对绝对引用且不改字符串表名函数(tmp_path):
    原 = tmp_path / '原.xlsx'
    表 = 表计划(名称='验证', 表=pd.DataFrame({'输入': [1, 2], '计算': ['=A2', '=A3']}))
    渲染簿计划(簿计划(表=[表]), str(原))
    文件 = tmp_path / '共享.xlsx'
    公式 = 'IF(A2=&quot;A2&quot;,SUM($B2:C$4,LOG10(A2)),\'A2表\'!D2)+SUM(A:C,2:4)'
    with ZipFile(原) as 读, ZipFile(文件, 'w', ZIP_DEFLATED) as 写:
        for 项 in 读.infolist():
            文 = 读.read(项.filename)
            if 项.filename == 'xl/worksheets/sheet1.xml':
                文 = re.sub(rb'<f>A2</f>', ('<f t="shared" si="0" ref="B2:B3">'+公式+'</f>').encode('utf-8'), 文)
                文 = 文.replace(b'<f>A3</f>', b'<f t="shared" si="0"/>')
            写.writestr(项, 文)
    with nullcontext(读簿(文件)) as 簿:
        assert 簿['验证'].值(3, 2) == '=IF(A3="A2",SUM($B3:C$4,LOG10(A3)),\'A2表\'!D3)+SUM(A:C,3:5)'


def test_真实Excel保存后全部公式仍可完整读回(tmp_path):
    from Excel验证 import 验证Excel报告
    文件 = tmp_path / 'Excel共享.xlsx'
    表 = 表计划(名称='验证', 表=pd.DataFrame({'输入': list(range(1, 101)), '计算': [f'=A{r}*2' for r in range(2, 102)]}))
    渲染簿计划(簿计划(表=[表]), str(文件))
    验证Excel报告(文件)
    with nullcontext(读簿(文件)) as 簿:
        assert [簿['验证'].值(r, 2) for r in range(2, 102)] == [f'=A{r}*2' for r in range(2, 102)]
