import pandas as pd
import pytest

from gui import _auto_mapping_for_columns
from input_precheck import _mapping_problems, _scope_values


def test_候选列换序不改变交易字段():
    columns = [
        "起息日期", "记账日期", "手续费金额", "可用余额", "付款方账号", "对方币种",
        "交易日期", "交易金额", "交易后余额", "账户", "币种",
    ]
    for labels in (columns, columns[::-1]):
        mapping = _auto_mapping_for_columns(labels, is_bank=True)
        assert {role: mapping[role] for role in ("date", "amount", "balance", "account", "currency")} == {
            "date": "交易日期", "amount": "交易金额", "balance": "交易后余额",
            "account": "账户", "currency": "币种",
        }


def test_账户列应参与核对范围取值():
    frame = pd.DataFrame({"日期": ["2026-01-02"], "金额": [100], "账户": ["0000123"]})
    assert _scope_values(frame, {"mode": "signed_amount", "date": "日期", "amount": "金额"}, "account")[1] == {"0000123"}


@pytest.mark.parametrize(("field", "wrong"), [
    ("account", "付款方账号"), ("currency", "对方币种"),
    ("date", "起息日期"), ("amount", "手续费金额"), ("balance", "可用余额"),
])
def test_已有映射须拦截已知错列(field, wrong):
    mapping = {"mode": "signed_amount", "date": "交易日期", "amount": "交易金额"}
    assert _mapping_problems({**mapping, field: wrong}, [*mapping.values(), wrong], "bank")


def test_交易时间单独一列时与交易日期分开():
    from 列识别 import recognize_columns

    frame = pd.DataFrame({
        "交易时间": ["13:02:00", "09:50:00"],
        "交易日期": ["2026-01-31", "2026-02-01"],
        "交易金额": [100, -40],
    })
    result = recognize_columns(frame, "bank")
    assert result.mapping["date"] == "交易日期"
    assert not result.choices


def test_从原表余额核实两个金额候选():
    from decimal import Decimal

    from data_loader import DataLoader
    from 列识别 import recognize_columns

    frame = pd.DataFrame({
        "交易日期": ["2026-01-31", "2026-02-01", "2026-02-02"],
        "交易金额(元)": [1, 2, 3],
        "交易金额（元）": [100, -40, 40],
        "交易后余额": [1100, 1060, 1100],
    })
    for source in (frame, frame.iloc[:, ::-1], frame.iloc[::-1]):
        result = recognize_columns(source, "bank")
        assert not result.choices and not result.problems
        assert result.mapping["amount"] == "交易金额（元）"
    records = DataLoader().standardize_data(frame, result.mapping, "bank")
    assert records["date"].dt.month.tolist() == [1, 2, 2]
    assert records["amount"].tolist() == [Decimal("100"), Decimal("-40"), Decimal("40")]


def test_单列分列并存时用数据判断金额格式():
    from 列识别 import recognize_columns

    frame = pd.DataFrame({
        "交易日期": ["2026-01-01", "2026-01-02"],
        "收入金额": [None, None], "支出金额": [None, None], "交易金额": [100, -40],
    })
    assert recognize_columns(frame, "bank").mapping["mode"] == "signed_amount"
    frame["收入金额"] = [100, None]
    frame["支出金额"] = [None, 40]
    assert recognize_columns(frame, "bank").mapping["mode"] == "debit_credit"
    frame["支出金额"] = [None, 41]
    assert recognize_columns(frame, "bank").choices["mode"] == ["debit_credit", "signed_amount"]
    selected = recognize_columns(frame, "bank", overrides={"mode": "signed_amount"})
    assert not selected.choices and selected.mapping["amount"] == "交易金额"


def test_币种由实际使用的金额列决定():
    from 列识别 import recognize_columns

    frame = pd.DataFrame({
        "凭证日期": ["2026-02-01"], "原币币种": ["USD"], "对方币种": ["EUR"],
        "本位币币种": ["CNY"], "本位币借方金额": [100], "本位币贷方金额": [None],
        "原币交易金额": [None],
    })
    result = recognize_columns(frame, "journal")
    assert result.mapping["amount_basis"] == "本位币"
    assert result.mapping["currency"] == "本位币币种"
    assert recognize_columns(frame.drop(columns="本位币币种"), "journal").mapping["currency"] is None


def test_只有时分秒的列不能当交易日期():
    from 列识别 import recognize_columns

    frame = pd.DataFrame({"交易时间": ["09:00:00"], "交易金额": [100]})
    assert recognize_columns(frame, "bank").problems == ["「交易时间」没有可解析的日期"]


def test_不同账户的两份表在逐笔匹配前拦截(tmp_path):
    from application import run_reconciliation
    from data_structures import MatcherConfig
    from input_precheck import InputPrecheckBlockedError

    bank, journal, output = (tmp_path / f"{name}.csv" for name in ("bank", "journal", "output"))
    for path, account in ((bank, "00000000101"), (journal, "00000000102")):
        pd.DataFrame({"日期": ["2026-01-01", "2026-01-02"], "交易金额": [100, -40],
                      "账户": [account, account]}).to_csv(path, index=False)
    matched = []
    with pytest.raises(InputPrecheckBlockedError, match="账户"):
        run_reconciliation(str(bank), str(journal), None, None, MatcherConfig(),
                           output_path=output.with_suffix(".xlsx"), matcher_ready=matched.append)
    assert not matched and not output.with_suffix(".xlsx").exists()


def test_数值一样但原币和本位币仍是不同候选():
    from 列识别 import recognize_columns

    frame = pd.DataFrame({
        "凭证日期": ["2026-01-02"], "原币币种": ["USD"], "本位币币种": ["CNY"],
        "原币借方金额": [100], "原币贷方金额": [0],
        "本位币借方金额": [100], "本位币贷方金额": [0],
    })
    result = recognize_columns(frame, "journal")
    assert set(result.choices["debit"]) == {"原币借方金额", "本位币借方金额"}
    picked = recognize_columns(frame, "journal", overrides={"debit": "原币借方金额"})
    assert not picked.choices
    assert picked.mapping["credit"] == "原币贷方金额" and picked.mapping["currency"] == "原币币种"
