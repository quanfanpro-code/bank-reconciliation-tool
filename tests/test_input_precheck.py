from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from data_loader import DataLoader


def _build_merged_header_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "合并表头.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["银行交易明细"])
    sheet.merge_cells("A1:E1")
    sheet.append(["日期", "发生额", None, "辅助信息", None])
    sheet.append([None, "借方", "贷方", "摘要", "对方户名"])
    sheet.merge_cells("A2:A3")
    sheet.merge_cells("B2:C2")
    sheet.merge_cells("D2:E2")
    sheet.append(["2026-01-01", 0, 100, "销售收款", "甲公司"])
    sheet.append(["2026-01-02", 20, 0, "支付手续费", "银行"])
    workbook.save(path)
    return path


def test_普通一行表头保持兼容(tmp_path):
    path = tmp_path / "普通表头.xlsx"
    pd.DataFrame(
        [{"日期": "2026-01-01", "收入": 100, "支出": 0, "摘要": "收款"}]
    ).to_excel(path, index=False)

    loader = DataLoader()
    structure = loader.detect_table_structure(path)
    data = loader.load_file(
        path,
        skiprows=structure.skiprows,
        header_rows=structure.header_rows,
        derived_columns=structure.columns,
    )

    assert structure.skiprows == 0
    assert structure.header_rows == 1
    assert structure.columns == ["日期", "收入", "支出", "摘要"]
    assert data.iloc[0].to_dict() == {
        "日期": "2026-01-01",
        "收入": 100,
        "支出": 0,
        "摘要": "收款",
    }


def test_标题与两行合并表头生成稳定无重复列名(tmp_path):
    path = _build_merged_header_workbook(tmp_path)

    loader = DataLoader()
    structure = loader.detect_table_structure(path)
    data = loader.load_file(
        path,
        skiprows=structure.skiprows,
        header_rows=structure.header_rows,
        derived_columns=structure.columns,
    )

    assert structure.skiprows == 1
    assert structure.header_rows == 2
    assert structure.columns == [
        "日期",
        "发生额｜借方",
        "发生额｜贷方",
        "辅助信息｜摘要",
        "辅助信息｜对方户名",
    ]
    assert len(structure.columns) == len(set(structure.columns))
    assert data.shape == (2, 5)
    assert data.iloc[0]["辅助信息｜对方户名"] == "甲公司"


def test_重复表头派生列名按出现顺序去重(tmp_path):
    path = tmp_path / "重复列名.csv"
    path.write_text(
        "日期,金额,金额,摘要\n2026-01-01,10,20,测试\n",
        encoding="utf-8-sig",
    )

    structure = DataLoader().detect_table_structure(path)

    assert structure.columns == ["日期", "金额", "金额_2", "摘要"]
    assert structure.columns == DataLoader().detect_table_structure(path).columns
