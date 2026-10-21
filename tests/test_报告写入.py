# -*- coding: utf-8 -*-
import pandas as pd
import pytest
from tests._excel假面 import load_workbook


def _小计划():
    from 排版计划 import 簿计划, 表计划, 列计划, 行计划, 格样式
    from 排版共享 import 条件规则

    表头样式 = 格样式(字体名="微软雅黑", 字号=10, 加粗=True, 填充色="FF1F4E79",
                    填充类型="solid", 水平="center", 垂直="center")
    df = pd.DataFrame({"名称": ["甲", "乙"], "金额": [1.5, "=B2*2"], "备注": ["批注格", None]})
    表 = 表计划(
        名称="测试表",
        表=df,
        表头样式=表头样式,
        列=[列计划(宽=20), 列计划(宽=12), 列计划(隐藏=True)],
        行={2: 行计划(高=30, outline=1), 3: 行计划(隐藏=True)},
        覆盖={(2, 2): 格样式(填充色="FFFFF2CC", 填充类型="solid", 数字格式="#,##0.00"),
              (2, 3): 格样式(批注=("提示", "核对"), 超链接="#测试表!A1")},
        验证=[{"type": "list", "formula1": '"是,否"', "sqref": "C2", "allow_blank": True}],
        条件格式=[("B2:B3", 条件规则(公式=["B2>1"], 填充色="FF0000"))],
        冻结="B2",
        筛选="A1:C3",
        tab色="FF0000",
        网格线=False,
        缩放=90,
        打印={"orientation": "landscape", "paperSize": 9, "fitToWidth": 1,
              "fitToHeight": 0, "fitToPage": True, "title_rows": "1:1",
              "左边距": 0.5, "右边距": 0.5},
    )
    隐藏表 = 表计划(名称="技术表", 表=pd.DataFrame({"x": [1]}), 状态="hidden")
    return 簿计划(表=[隐藏表, 表], active=1, 自动重算=True)


def test_渲染簿计划(tmp_path):
    from 报告写入 import 渲染簿计划

    输出 = tmp_path / "出.xlsx"
    渲染簿计划(_小计划(), str(输出))
    wb = load_workbook(输出)
    assert wb.sheetnames == ["技术表", "测试表"]
    assert wb.active == wb["测试表"]
    assert wb["技术表"].sheet_state == "hidden"
    assert wb.calculation.fullCalcOnLoad

    ws = wb["测试表"]
    assert ws.sheet_properties.tabColor.rgb == "00FF0000" or ws.sheet_properties.tabColor.rgb == "FF0000"
    assert ws.freeze_panes == "B2"
    assert ws.auto_filter.ref == "A1:C3"
    assert ws.sheet_view.showGridLines is False
    assert ws.sheet_view.zoomScale == 90

    # 表头样式
    h = ws["A1"]
    assert h.value == "名称"
    assert h.font.name == "微软雅黑" and h.font.bold
    assert h.fill.fill_type == "solid" and h.fill.fgColor.rgb == "FF1F4E79"
    assert h.alignment.horizontal == "center"

    # 值与公式
    assert ws["A2"].value == "甲"
    assert ws["B2"].value == 1.5
    assert ws["B3"].value == "=B2*2"

    # 列
    assert ws.column_dimensions["A"].width == 20
    assert ws.column_dimensions["C"].hidden

    # 行
    assert ws.row_dimensions[2].height == 30
    assert ws.row_dimensions[2].outlineLevel == 1
    assert ws.row_dimensions[3].hidden

    # 格覆盖
    b2 = ws["B2"]
    assert b2.fill.fgColor.rgb == "FFFFF2CC"
    assert b2.number_format == "#,##0.00"
    c2 = ws["C2"]
    assert c2.comment.text == "提示"
    assert c2.hyperlink.target == "#测试表!A1"

    # 数据验证与条件格式
    dvs = list(ws.data_validations.dataValidation)
    assert len(dvs) == 1 and dvs[0].type == "list" and "C2" in str(dvs[0].sqref)
    assert len(ws.conditional_formatting._cf_rules) == 1

    # 打印
    assert ws.page_setup.orientation == "landscape"
    assert ws.page_setup.fitToWidth == 1
    assert ws.print_title_rows == "$1:$1"
    assert ws.page_margins.left == 0.5


def test_样式覆盖优先级_盖优于列(tmp_path):
    from 排版计划 import 簿计划, 表计划, 列计划, 格样式
    from 报告写入 import 渲染簿计划

    表 = 表计划(
        名称="t",
        表=pd.DataFrame({"a": [1]}),
        列=[列计划(基础样式=格样式(填充色="FF0000FF", 填充类型="solid", 加粗=True))],
        覆盖={(2, 1): 格样式(填充色="FF00FF00", 填充类型="solid")},
    )
    输出 = tmp_path / "出.xlsx"
    渲染簿计划(簿计划(表=[表]), str(输出))
    c = load_workbook(输出)["t"]["A2"]
    assert c.fill.fgColor.rgb == "FF00FF00"   # 覆盖的填充胜出
    assert c.font.bold                          # 覆盖未设加粗，列基础样式保留


def test_新建簿计划_默认裸表(tmp_path):
    from 排版计划 import 新建簿计划
    from 报告写入 import 渲染簿计划

    计划 = 新建簿计划({"甲表": pd.DataFrame({"x": [1, 2]}), "乙表": pd.DataFrame({"y": [3]})})
    输出 = tmp_path / "出.xlsx"
    渲染簿计划(计划, str(输出))
    wb = load_workbook(输出)
    assert wb.sheetnames == ["甲表", "乙表"]
    assert wb["甲表"]["A2"].value == 1


def test_大表快速通道_列默认样式(tmp_path, monkeypatch):
    """大表（> 快速通道行数）数据格只写值、整列样式落到列维度；特例格仍逐格。"""
    import 排版计划
    from 排版计划 import 格样式, 簿计划, 列计划, 行计划, 表计划
    from 报告写入 import 渲染簿计划

    monkeypatch.setattr(排版计划, "_快速通道行数", 3)
    df = pd.DataFrame({
        "文本": ["甲", "乙", "丙", "丁"],
        "金额": [1000.0, 2000.5, 3000.0, 4000.0],
        "链接": ["查看", None, None, None],
    })
    表 = 表计划(
        名称="大表",
        表=df,
        表头样式=格样式(字体名="微软雅黑", 字号=10, 加粗=True,
                     填充色="FF1F4E79", 填充类型="solid"),
        列=[列计划(宽=20), 列计划(宽=12), 列计划(宽=10)],
        # 行样式行仍走完整逐格路径（行样式加粗/红字在格变换后保留，填充被清空——
        # 与旧版逐格行为一致：紧凑变换对整行清填充）
        行={3: 行计划(样式=格样式(加粗=True, 字色="FF0000"))},
        # 覆盖格（超链接）仍走完整逐格路径
        覆盖={(2, 3): 格样式(超链接="#大表!A1")},
    )
    表.格变换 = 排版计划._做变换(False)
    输出 = tmp_path / "出.xlsx"
    渲染簿计划(簿计划(表=[表]), str(输出))
    wb = load_workbook(输出)
    ws = wb["大表"]
    # 普通数据格：无逐格样式，列默认兜底（字体经列样式继承读回；数字格式见下方 XML 断言）
    assert ws["B4"].value == 3000.0
    assert ws["B4"].font.name == "MiSans" and ws["B4"].font.size == 9
    # 列默认样式带格变换效果：MiSans 9、文本列左对齐（单元格继承读回）
    assert ws["A2"].font.name == "MiSans" and ws["A2"].font.size == 9
    assert ws["A2"].alignment.horizontal == "left"
    # 金额列的千分位/右对齐经列样式承载：openpyxl 读不回，直接断言 XML
    import re
    import zipfile
    zp = zipfile.ZipFile(输出)
    表 = zp.read("xl/worksheets/sheet1.xml").decode("utf-8")
    样 = zp.read("xl/styles.xml").decode("utf-8")
    列格 = re.search(r'<col min="2" max="2"[^>]*style="(\d+)"', 表)
    assert 列格, "金额列应带列样式"
    xf段 = re.search(r"<cellXfs count=\"\d+\">(.*?)</cellXfs>", 样, re.S).group(1)
    xf = re.findall(r"<xf [^>]*>", xf段)[int(列格.group(1))]
    assert 'applyNumberFormat="1"' in xf and 'applyAlignment="1"' in xf
    assert "#,##0.00;[Red]-#,##0.00" in 样 and 'horizontal="right"' in 样
    # 表头仍是显式表头样式
    assert ws["A1"].font.name == "微软雅黑" and ws["A1"].font.bold
    # 覆盖格（超链接）走完整路径：MiSans 字体 + 超链接生效
    assert ws["C2"].hyperlink is not None
    assert ws["C2"].font.name == "MiSans"
    # 行样式行走完整路径：加粗红字保留、填充被紧凑变换清空（与旧版逐格一致）
    assert ws["B3"].font.bold
    assert ws["B3"].font.color.rgb[-6:] == "FF0000"
    assert ws["B3"].font.name == "MiSans"
    assert ws["B3"].fill.fill_type is None
