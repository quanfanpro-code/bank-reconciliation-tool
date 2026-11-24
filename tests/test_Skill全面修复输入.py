from datetime import datetime
import pytest
import pandas as pd
from openpyxl import load_workbook
from xlsxwriter import Workbook

from data_loader import DataLoader
from data_structures import MatcherConfig
from 运行核对 import SkillMatcher
from 银行核对入口 import AgentReview, pair_inputs, run_accounts


def 原表(路径, 工作表, *, 日期系统1904=False):
    with Workbook(路径, {"date_1904": 日期系统1904}) as 簿:
        for 名称, 行列 in 工作表.items():
            页 = 簿.add_worksheet(名称)
            日期格式 = 簿.add_format({"num_format": "yyyy-mm-dd"})
            for 序号, 行 in enumerate(行列):
                for 列号, 值 in enumerate(行):
                    页.write(序号, 列号, 值, 日期格式 if isinstance(值, datetime) else None)


表头 = ["交易日期", "交易金额", "本方账号", "币种", "摘要"]
已对上 = ["2025-05-12", 150, "001122", "CNY", "销售款"]


def 跑一次(tmp_path, 流水, 日记账):
    return run_accounts(流水, 日记账, tmp_path / "报告", MatcherConfig(), tmp_path / "会话",
                        matcher_factory=SkillMatcher, log=lambda _: None)


def 验证月表(结果, 列名, 数值):
    assert not 结果["待处理"], 结果
    报告, = 结果["完成"]
    with open(报告["报告"], "rb") as 文件:
        簿 = load_workbook(文件, read_only=True, data_only=True)
        try:
            表 = list(簿["月度核对"].values)
            assert 表[1][表[0].index(列名)] == 数值
            return list(簿["交易明细"].values)
        finally:
            簿.close()


def test_不同页签的明细都进入同一个账户(tmp_path):
    流水, 日记账 = tmp_path / "流水.xlsx", tmp_path / "日记账.xlsx"
    原表(流水, {"汇总": [["月份", "收入合计", "支出合计", "交易笔数"], ["2025-05", 150, 90000, 2]], "本册上页": [表头, 已对上],
               "本册后页": [表头, ["2025-05-26", -90000, "001122", "CNY", "设备款"]]})
    原表(日记账, {"账": [表头, 已对上]})
    明细 = 验证月表(跑一次(tmp_path, 流水, 日记账), "支出金额差额", 90000)
    assert any("本册后页" in str(值) and "第2行" in str(值) for 行 in 明细 for 值 in 行)


def test_无规律文件名仍按表内账号配户(tmp_path):
    流水, 日记账 = tmp_path / "流水", tmp_path / "日记账"
    流水.mkdir()
    日记账.mkdir()
    a, b = 流水 / "明细.xlsx", 日记账 / "五月.xlsx"
    for 文件 in (a, b):
        原表(文件, {"明细": [表头, 已对上]})
    assert pair_inputs(流水, 日记账) == ([(a, b)], [])


def test_逐户出表且文件名里的旧账号不覆盖原表(tmp_path):
    流水, 日记账 = tmp_path / "流水", tmp_path / "日记账"
    流水.mkdir()
    日记账.mkdir()
    原表(流水 / "甲公司-甲银行-009999.xlsx", {"明细": [表头, 已对上]})
    原表(日记账 / "五月.xlsx", {"明细": [表头, 已对上]})
    原表(流水 / "乙户.xlsx", {"明细": [表头, ["2025-05-31", -45, "002233", "CNY", "手续费"]]})
    结果 = 跑一次(tmp_path, 流水, 日记账)
    assert not 结果["待处理"], 结果
    assert len(结果["完成"]) == 2
    for 报告 in 结果["完成"]:
        with open(报告["报告"], "rb") as 文件:
            簿 = load_workbook(文件, read_only=True, data_only=True)
            try:
                总览 = dict(row[:2] for row in 簿["核对总览"].values if row[0])
                assert "009999" not in 总览["银行及账号"]
                if 总览["银行及账号"] == "002233":
                    assert "缺少日记账" in 总览["月度核对结果"]
                    表头行, 本月 = list(簿["月度核对"].values)
                    assert 本月[表头行.index("银行支出金额")] == 45
                    assert 本月[表头行.index("支出金额差额")] is None
            finally:
                簿.close()


def test_按工作簿的日期系统读取(tmp_path):
    流水 = tmp_path / "流水.xlsx"
    原表(流水, {"明细": [表头, [datetime(2025, 5, 20), *已对上[1:]]]}, 日期系统1904=True)
    assert pd.Timestamp(DataLoader().load_file(流水).iloc[0, 0]) == pd.Timestamp("2025-05-20")


def test_csv空行后的原行号和空日期保留(tmp_path):
    流水 = tmp_path / "流水.csv"
    流水.write_text("日期,交易金额,摘要\n2025-05-20,100,一笔收款\n\n,200,另一笔", encoding="utf-8-sig")
    加载 = DataLoader()
    原始 = 加载.load_file(流水)
    assert 原始["__file_row__"].tolist() == [2, 3, 4]
    明细 = 加载.standardize_data(原始, {"date": "日期", "amount": "交易金额", "mode": "signed_amount"}, "bank")
    assert list(明细["original_file_row"]) == [2]


@pytest.mark.parametrize("日期列,金额列", [("资金变动之日", "实际增减值"), ("收付发生于", "现金流量数")])
def test_表头问题由同次任务的Agent处理(tmp_path, monkeypatch, 日期列, 金额列):
    流水, 日记账 = tmp_path / "流水.xlsx", tmp_path / "日记账.xlsx"
    原表(流水, {"明细": [[金额列, 日期列, *表头[2:]], [已对上[1], 已对上[0], *已对上[2:]]]})
    原表(日记账, {"明细": [表头, 已对上]})
    def 回复(_self, packet, **_kwargs):
        assert packet["kind"] == "columns"
        assert 日期列 in packet["preview"][0] and 金额列 in packet["preview"][0]
        return {"columns": {"bank": {"date": 日期列, "amount": 金额列}}}, ""
    monkeypatch.setattr(AgentReview, "exchange", 回复)
    验证月表(跑一次(tmp_path, 流水, 日记账), "收入金额差额", 0)


def test_读取银行文件明示的零交易(tmp_path):
    流水, 日记账 = tmp_path / "流水.xlsx", tmp_path / "日记账.xlsx"
    期间 = ["查询起止：2025-05-01 至 2025-05-31  本方账号：001122  币种：CNY"]
    原表(流水, {"明细": [期间, ["交易笔数 0，收入 0，支出 0，2025-05-31 余额 1000"], 表头]})
    原表(日记账, {"明细": [期间, 表头, ["2025-05-31", 100, "001122", "CNY", "回款"]]})
    验证月表(跑一次(tmp_path, 流水, 日记账), "收入金额差额", -100)
