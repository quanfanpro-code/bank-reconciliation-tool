# -*- coding: utf-8 -*-
"""排版对比工具的行为测试：一致返回空，差异必须报得出。"""
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from 排版对比 import 对比工作簿


def _存(tmp_path, name, wb):
    wb.save(tmp_path / name)
    return str(tmp_path / name)


def test_一致文件返回空(tmp_path):
    for name in ("a.xlsx", "b.xlsx"):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "x"
        ws.column_dimensions["A"].width = 14
        wb.save(tmp_path / name)
    assert 对比工作簿(str(tmp_path / "a.xlsx"), str(tmp_path / "b.xlsx")) == []


def test_能报出行高差异(tmp_path):
    wb = Workbook()
    wb.active["A1"] = "x"
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    wb.active["A1"] = "x"
    wb.active.row_dimensions[1].height = 33
    b = _存(tmp_path, "b.xlsx", wb)
    差异 = 对比工作簿(a, b)
    assert any("行高" in d for d in 差异), f"未报出行高差异: {差异}"


def test_能报出值与填充差异(tmp_path):
    wb = Workbook()
    wb.active["A1"] = "x"
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    wb.active["A1"] = "y"
    b = _存(tmp_path, "b.xlsx", wb)
    差异 = 对比工作簿(a, b)
    assert any("值" in d for d in 差异), f"未报出值差异: {差异}"


def test_能报出表顺序差异(tmp_path):
    wb = Workbook()
    wb.create_sheet("甲")
    wb.create_sheet("乙")
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    wb.create_sheet("乙")
    wb.create_sheet("甲")
    b = _存(tmp_path, "b.xlsx", wb)
    差异 = 对比工作簿(a, b)
    assert any("顺序" in d or "工作表" in d for d in 差异), f"未报出表顺序差异: {差异}"


def test_能报出字体与数字格式差异(tmp_path):
    wb = Workbook()
    wb.active["A1"] = 0.5
    wb.active["A1"].font = Font(name="MiSans", size=9, bold=True)
    wb.active["A1"].number_format = "0.00%"
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    wb.active["A1"] = 0.5
    wb.active["A1"].font = Font(name="Arial", size=11)
    wb.active["A1"].number_format = "#,##0.00"
    b = _存(tmp_path, "b.xlsx", wb)
    差异 = 对比工作簿(a, b)
    assert any("字体" in d for d in 差异), f"未报出字体差异: {差异}"
    assert any("数字格式" in d for d in 差异), f"未报出数字格式差异: {差异}"


def test_能报出数据验证与条件格式差异(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    dv = DataValidation(type="list", formula1='"甲,乙"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("A1")
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "x"
    b = _存(tmp_path, "b.xlsx", wb)
    差异 = 对比工作簿(a, b)
    assert any("数据验证" in d for d in 差异), f"未报出数据验证差异: {差异}"

    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws.conditional_formatting.add(
        "A1:A9",
        FormulaRule(formula=["A1<>0"], fill=PatternFill("solid", fgColor="FFF2CC")),
    )
    c = _存(tmp_path, "c.xlsx", wb)
    差异 = 对比工作簿(a, c)
    assert any("条件格式" in d for d in 差异), f"未报出条件格式差异: {差异}"


def test_能报出超链接与outline差异(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "去"
    ws["A1"].hyperlink = "#'乙'!A1"
    ws.row_dimensions[2].outlineLevel = 1
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    wb.active["A1"] = "去"
    b = _存(tmp_path, "b.xlsx", wb)
    差异 = 对比工作簿(a, b)
    assert any("超链接" in d for d in 差异), f"未报出超链接差异: {差异}"
    assert any("outline" in d.lower() or "分级" in d for d in 差异), f"未报出outline差异: {差异}"


def test_默认值与缺失等价不报差异(tmp_path):
    # 显式默认（隐藏=False、outline=0）与完全未设置应视为一致
    wb = Workbook()
    wb.active["A1"] = "x"
    wb.active.row_dimensions[2].hidden = False
    wb.active.row_dimensions[2].outlineLevel = 0
    a = _存(tmp_path, "a.xlsx", wb)
    wb = Workbook()
    wb.active["A1"] = "x"
    b = _存(tmp_path, "b.xlsx", wb)
    assert 对比工作簿(a, b) == []
