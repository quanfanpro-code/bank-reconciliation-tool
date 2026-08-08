from pathlib import Path
from decimal import Decimal

import pandas as pd
from openpyxl import Workbook

import input_precheck
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


def _structure(columns, *, ambiguous=False, candidates=()):
    return input_precheck.TableStructure(
        skiprows=0,
        header_rows=1,
        columns=list(columns),
        score=10,
        ambiguous=ambiguous,
        candidates=tuple(candidates),
        explanation="测试结构",
    )


def _standardized(rows, source):
    return pd.DataFrame(
        [
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "summary": summary,
                "source": source,
            }
            for date, amount, summary in rows
        ]
    )


def _signed_mapping():
    return {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "auxiliary_text_columns": ["摘要", "对方户名"],
        "mode": "signed_amount",
    }


def _build_report(
    raw_bank,
    raw_journal,
    bank,
    journal,
    *,
    bank_mapping=None,
    journal_mapping=None,
    bank_structure=None,
    journal_structure=None,
    parse_errors=(),
):
    mapping = _signed_mapping()
    return input_precheck.build_input_precheck(
        raw_bank=raw_bank,
        raw_journal=raw_journal,
        bank=bank,
        journal=journal,
        bank_mapping=bank_mapping or mapping,
        journal_mapping=journal_mapping or mapping,
        bank_structure=bank_structure or _structure(raw_bank.columns),
        journal_structure=journal_structure or _structure(raw_journal.columns),
        parse_errors=list(parse_errors),
    )


def test_日期范围与收支合计分别比较且只提示():
    raw_bank = pd.DataFrame(
        [
            {"日期": "2026-01-01", "金额": 100, "摘要": "收款", "对方户名": "甲"},
            {"日期": "2026-01-03", "金额": -30, "摘要": "付款", "对方户名": "乙"},
        ]
    )
    raw_journal = pd.DataFrame(
        [
            {"日期": "2026-01-02", "金额": 80, "摘要": "收款", "对方户名": "甲"},
            {"日期": "2026-01-04", "金额": -20, "摘要": "付款", "对方户名": "乙"},
        ]
    )

    report = _build_report(
        raw_bank,
        raw_journal,
        _standardized(
            [("2026-01-01", 100, "收款"), ("2026-01-03", -30, "付款")],
            "bank",
        ),
        _standardized(
            [("2026-01-02", 80, "收款"), ("2026-01-04", -20, "付款")],
            "journal",
        ),
    )

    date_item = next(item for item in report.items if item.name == "日期范围")
    amount_item = next(item for item in report.items if item.name == "金额合计")
    assert report.has_blockers is False
    assert date_item.status == "提示"
    assert amount_item.status == "提示"
    assert amount_item.bank_result == "收入 100.00；支出 30.00"
    assert amount_item.journal_result == "收入 80.00；支出 20.00"
    assert amount_item.comparison == "收入差额 20.00；支出差额 10.00"


def test_无效方向与必填列缺失会阻止运行():
    raw = pd.DataFrame(
        [{"日期": "2026-01-01", "金额": 100, "摘要": "测试", "对方户名": "甲"}]
    )
    mapping = {
        "date": "日期",
        "amount": "金额",
        "direction": "方向",
        "summary": "摘要",
        "mode": "single_amount_with_direction",
    }
    report = _build_report(
        raw,
        raw,
        _standardized([("2026-01-01", 100, "测试")], "bank"),
        _standardized([("2026-01-01", 100, "测试")], "journal"),
        bank_mapping=mapping,
        journal_mapping=mapping,
        parse_errors=[
            {
                "type": "方向解析失败",
                "source_type": "bank",
                "row": 2,
                "original_value": "未知",
                "column": "方向",
            }
        ],
    )

    assert report.has_blockers is True
    assert next(item for item in report.items if item.name == "金额方向").status == "阻止"
    assert next(item for item in report.items if item.name == "必填字段").status == "阻止"
    assert "方向" in report.blocker_message()


def test_少量解析异常非交易行和低文字非空率只提示():
    raw_bank = pd.DataFrame(
        [
            {"日期": "2026-01-01", "金额": 100, "摘要": "", "对方户名": ""},
            {"日期": None, "金额": None, "摘要": None, "对方户名": None},
            {"日期": None, "金额": 100, "摘要": "本期合计", "对方户名": None},
            {"日期": "日期", "金额": "金额", "摘要": "摘要", "对方户名": "对方户名"},
        ]
    )
    raw_journal = pd.DataFrame(
        [{"日期": "2026-01-01", "金额": 100, "摘要": "测试", "对方户名": "甲"}]
    )
    report = _build_report(
        raw_bank,
        raw_journal,
        _standardized([("2026-01-01", 100, "")], "bank"),
        _standardized([("2026-01-01", 100, "测试")], "journal"),
        parse_errors=[
            {
                "type": "日期解析失败",
                "source_type": "bank",
                "row": 5,
                "original_value": "日期错误",
                "column": "日期",
            }
        ],
    )

    assert report.has_blockers is False
    assert next(item for item in report.items if item.name == "非交易行").status == "提示"
    assert next(item for item in report.items if item.name == "辅助文字完整性").status == "提示"
    assert "少量日期或金额解析失败" in report.warning_message()


def test_日期或金额全部无可用记录会阻止():
    raw = pd.DataFrame([{"日期": "错误", "金额": "错误", "摘要": "测试", "对方户名": "甲"}])
    empty = pd.DataFrame(columns=["date", "amount", "summary", "source"])

    report = _build_report(raw, raw, empty, empty)

    assert report.has_blockers is True
    assert next(item for item in report.items if item.name == "日期范围").status == "阻止"
    assert next(item for item in report.items if item.name == "金额合计").status == "阻止"


def test_检查结果可直接转换为固定列报告表():
    raw = pd.DataFrame(
        [{"日期": "2026-01-01", "金额": 100, "摘要": "测试", "对方户名": "甲"}]
    )
    report = _build_report(
        raw,
        raw,
        _standardized([("2026-01-01", 100, "测试")], "bank"),
        _standardized([("2026-01-01", 100, "测试")], "journal"),
    )

    table = report.to_dataframe()

    assert list(table.columns) == [
        "检查项目",
        "银行流水结果",
        "银行日记账结果",
        "双方比较结果",
        "状态",
        "说明",
    ]
    assert table["检查项目"].tolist() == [
        "文件读取",
        "表格结构",
        "日期范围",
        "金额方向",
        "金额合计",
        "非交易行",
        "必填字段",
        "辅助文字完整性",
    ]
