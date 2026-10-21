import hashlib
import json
import subprocess
from pathlib import Path

from tests._excel假面 import load_workbook

from 输入模板 import generate_input_templates
from 项目记录 import LocalProjectStore


def _sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def test_列映射模板按表头指纹保存和复用(tmp_path):
    store = LocalProjectStore(tmp_path / "records")
    mapping = {"date": "交易日期", "amount": "金额", "summary": "摘要"}

    fingerprint = store.save_mapping_template("bank", ["金额", "摘要", "交易日期"], mapping, "建行流水")
    loaded = store.load_mapping_template("bank", ["交易日期", "金额", "摘要"])

    assert loaded["fingerprint"] == fingerprint
    assert loaded["mapping"] == mapping
    assert loaded["name"] == "建行流水"


def test_项目记录保存输入与报告指纹并可重新打开(tmp_path):
    bank = tmp_path / "bank.xlsx"
    journal = tmp_path / "journal.xlsx"
    report = tmp_path / "report.xlsx"
    bank.write_bytes(b"bank")
    journal.write_bytes(b"journal")
    report.write_bytes(b"report")
    store = LocalProjectStore(tmp_path / "records")

    record = store.save_project(
        bank_path=bank,
        journal_path=journal,
        report_path=report,
        bank_mapping={"date": "日期"},
        journal_mapping={"date": "凭证日期"},
        parameters={"tolerance_days": 3},
        result_counts={"selected_groups": 5},
    )
    latest = store.list_projects()[0]

    assert latest["project_id"] == record["project_id"]
    assert latest["files"]["report"]["sha256"] == _sha256(report)
    assert store.resolve_report(record["project_id"]) == report


def test_输入模板生成银行流水和序时账两个可直接填写文件(tmp_path):
    paths = generate_input_templates(tmp_path)

    assert {path.name for path in paths} == {"银行流水导入模板.xlsx", "银行存款序时账导入模板.xlsx"}
    bank_book = load_workbook(tmp_path / "银行流水导入模板.xlsx", read_only=True)
    journal_book = load_workbook(tmp_path / "银行存款序时账导入模板.xlsx", read_only=True)
    assert bank_book.active.cell(1, 1).value == "交易日期"
    assert journal_book.active.cell(1, 1).value == "凭证日期"
    assert {"交易日期", "收入金额", "支出金额", "账户", "币种", "流水号", "摘要"}.issubset(
        {cell.value for cell in bank_book.active[1]}
    )
    assert {"凭证日期", "借方金额", "贷方金额", "账户", "币种", "凭证号", "摘要"}.issubset(
        {cell.value for cell in journal_book.active[1]}
    )


def test_windows双击入口提供自检且无需手工输入命令():
    root = Path(__file__).resolve().parents[1]
    launcher = root / "启动银行流水核对工具.cmd"

    result = subprocess.run(["cmd.exe", "/c", str(launcher), "--check"], cwd=root, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=20)

    assert result.returncode == 0
    assert "READY" in result.stdout


def test_主程序提供打包后可用的自检入口():
    root = Path(__file__).resolve().parents[1]
    result = subprocess.run(
        [
            r"C:\Users\27651\AppData\Local\Programs\Python\Python314\python.exe",
            str(root / "main.py"),
            "--check",
        ],
        cwd=root,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=20,
    )

    assert result.returncode == 0
    assert "READY" in result.stdout
