"""第二阶段独立反例：完整凭证、交易级流水号和批次差异口径。"""

from decimal import Decimal

import pandas as pd

from data_loader import DataLoader, ParseErrorCollector
from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter
from gui import _auto_mapping_for_columns
from 业务分组 import complete_groups, row_business


def 原始(金额们, 来源="journal", 摘要="转账", 对方=""):
    行们 = []
    for 序号, 金额 in enumerate(金额们, start=1):
        借方 = 金额 if 来源 == "journal" else 0
        贷方 = 0 if 来源 == "journal" else 金额
        行们.append({
            "日期": "2026-08-10",
            "摘要": 摘要,
            "对方户名": 对方,
            "借方金额": 借方,
            "贷方金额": 贷方,
            "凭证字": "",
            "凭证号": "",
            "交易流水号": "",
            "余额": "",
            "原行": 序号,
        })
    return pd.DataFrame(行们)


def 映射():
    return {
        "date": "日期",
        "summary": "摘要",
        "mode": "debit_credit",
        "debit": "借方金额",
        "credit": "贷方金额",
        "voucher_word": "凭证字",
        "voucher": "凭证号",
        "balance": "余额",
        "amount_basis": "本位币",
        "auxiliary_text_columns": ["摘要", "对方户名", "交易流水号"],
    }


def 核对(银行, 账, **配置):
    加载器 = DataLoader()
    实例 = Matcher(
        加载器.standardize_data(银行, 映射(), "bank"),
        加载器.standardize_data(账, 映射(), "journal"),
        MatcherConfig(**配置),
    )
    实例.run()
    return 实例


def test_完整凭证组不可被组合搜索切成其中两行自动确认():
    银行 = 原始([1000], "bank", "银行转账")
    账 = 原始([400, 600, 500], 摘要="材料")
    账["摘要"] = ["材料", "人工", "运费"]
    账["凭证字"] = "记"
    账["凭证号"] = "001"

    实例 = 核对(银行, 账)

    assert not any(
        set(候选.journal_idxs) in ({0}, {1}, {2}, {0, 1}, {0, 2}, {1, 2})
        for 候选 in 实例.selected_candidates
    ), "同一完整凭证必须作为原子组，不能为了凑1000元只取400+600"
    assert any(
        候选.bank_idxs == (0,) and 候选.journal_idxs == (0, 1, 2)
        and 候选.metrics.total_diff_li == PrecisionEngine.to_integer_li(500)
        and 候选.processing_status is ProcessingStatus.AUTO_CLASSIFIED
        for 候选 in 实例.selected_candidates
    ), "整张凭证应以1500对1000保留500元差异"


def test_交易级流水号各不相同不拆散35笔工资对一笔汇总():
    银行 = 原始([100] * 35, "bank", "8月工资")
    银行["交易流水号"] = [f"B{i:03d}" for i in range(1, 36)]
    账 = 原始([3500], 摘要="8月工资汇总")
    账["交易流水号"] = "J001"

    实例 = 核对(银行, 账)

    assert len(实例.selected_candidates) == 1
    候选 = 实例.selected_candidates[0]
    assert 候选.bank_idxs == tuple(range(35)) and 候选.journal_idxs == (0,)
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
    assert "交易流水号" not in 候选.processing_reason
    assert not 候选.evidence.get("business_conflicts")


def test_同月相同凭证号但凭证字不同必须形成两个原子组():
    账 = 原始([400, 600, 300, 700])
    账["凭证字"] = ["记", "记", "付", "付"]
    账["凭证号"] = "001"
    标准账 = DataLoader().standardize_data(账, 映射(), "journal")
    业务行 = {int(i): row_business(行) for i, 行 in 标准账.iterrows()}

    凭证组 = [组 for 组 in complete_groups(业务行, 31) if 组[0][0] == "凭证"]

    assert len(凭证组) == 2
    assert {组[0][2:4] for 组 in 凭证组} == {("记", "001"), ("付", "001")}
    按凭证字金额 = {
        组[0][2]: sorted(int(标准账.at[索引, "amount_decimal"]) for 索引 in 组[1])
        for 组 in 凭证组
    }
    assert 按凭证字金额 == {
        "记": [PrecisionEngine.to_integer_li(400), PrecisionEngine.to_integer_li(600)],
        "付": [PrecisionEngine.to_integer_li(300), PrecisionEngine.to_integer_li(700)],
    }


def test_工资明细对汇总的差异线索不得把层级笔数差说成缺多少笔():
    银行 = 原始([100] * 34, "bank", "8月工资")
    账 = 原始([3500], 摘要="8月工资汇总")

    实例 = 核对(银行, 账)

    候选 = 实例.selected_candidates[0]
    线索 = 候选.evidence.get("batch_review_hint", "")
    assert "34笔明细" in 线索 and "1笔汇总" in 线索
    assert "少33笔" not in 线索 and "多33笔" not in 线索
    assert "相当于1笔" in 线索 and "不能确定" in 线索


def test_余额解析失败必须留存原行字段及原值():
    银行 = 原始([100], "bank")
    银行.loc[0, "余额"] = "余额坏值"
    收集器 = ParseErrorCollector()

    DataLoader(error_collector=收集器).standardize_data(银行, 映射(), "bank")

    错误 = [项目 for 项目 in 收集器.get_all_errors() if 项目["type"] == "余额解析失败"]
    assert len(错误) == 1
    assert 错误[0]["column"] == "余额"
    assert 错误[0]["original_value"] == "余额坏值"
    assert 错误[0]["row"] == 2


def test_实际界面自动映射能区分本方账号余额凭证字和凭证号():
    结果 = _auto_mapping_for_columns(
        ["交易日期", "本方账号", "账户余额", "原币币种", "凭证字", "凭证号", "本位币借方金额", "本位币贷方金额"],
        is_bank=False,
    )

    assert 结果["account"] == "本方账号"
    assert 结果["balance"] == "账户余额"
    assert 结果["currency"] == "原币币种"
    assert 结果["voucher_word"] == "凭证字"
    assert 结果["voucher"] == "凭证号"
    assert 结果["amount_basis"] == "本位币"


def test_匹配组成报告逐行保留凭证字与凭证号证据():
    银行 = 原始([1000], "bank", "转账")
    账 = 原始([400, 600], 摘要="项目明细")
    账["凭证字"] = "记"
    账["凭证号"] = "001"
    实例 = 核对(银行, 账)

    组成 = Reporter(实例).build_report_tables(实例.config)["匹配组成"]
    日记账行 = 组成.loc[组成["来源"] == "日记账"]

    assert set(日记账行["凭证字"]) == {"记"}
    assert set(日记账行["凭证号"].astype(str)) == {"001"}
    assert set(日记账行["原凭证字列名"]) == {"凭证字"}
    assert set(日记账行["解析后凭证字"]) == {"记"}


def test_共同批次号内不同订单号不得拆散供应商批付():
    银行 = 原始([100, 200, 300, 400], "bank", "供应商批付", "甲公司")
    银行["批次号"] = "B001"
    银行["订单号"] = ["O1", "O2", "O3", "O4"]
    账 = 原始([1000], 摘要="供应商批付汇总", 对方="甲公司")
    账["批次号"] = "B001"
    账["订单号"] = ""
    列映射 = 映射()
    列映射["auxiliary_text_columns"] = [
        "摘要", "对方户名", "批次号", "订单号"
    ]
    加载器 = DataLoader()
    实例 = Matcher(
        加载器.standardize_data(银行, 列映射, "bank"),
        加载器.standardize_data(账, 列映射, "journal"),
        MatcherConfig(),
    )
    实例.run()

    assert len(实例.selected_candidates) == 1
    候选 = 实例.selected_candidates[0]
    assert 候选.bank_idxs == (0, 1, 2, 3) and 候选.journal_idxs == (0,)
    assert 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
    assert 候选.evidence.get("shared_business_id") is True
    assert not 候选.evidence.get("business_conflicts")


def test_同日同额但无任何业务文字或编号不得自动确认():
    银行 = 原始([100], "bank", "", "")
    账 = 原始([100], 摘要="", 对方="")

    实例 = 核对(银行, 账)

    assert 实例.selected_candidates
    assert 实例.selected_candidates[0].processing_status is ProcessingStatus.FLAGGED
    assert "业务依据不足" in 实例.selected_candidates[0].processing_reason


def test_自动确认最低综合可信度配置必须实际参与分流():
    银行 = 原始([100], "bank", "设备采购", "")
    账 = 原始([100], 摘要="设备采购", 对方="")

    实例 = 核对(银行, 账, auto_confirm_score=100)

    候选 = 实例.selected_candidates[0]
    assert 候选.scores.total < 100
    assert 候选.processing_status is ProcessingStatus.FLAGGED
    assert "低于自动确认门槛" in 候选.processing_reason


def test_选择本位币计算后仍保留原币本位币账户和币种原值():
    账 = pd.DataFrame([{
        "日期": "2026-08-10",
        "摘要": "设备采购",
        "本方账号": "A001",
        "原币币种": "USD",
        "本位币币种": "CNY",
        "原币借方金额": 100,
        "原币贷方金额": 0,
        "本位币借方金额": 720,
        "本位币贷方金额": 0,
    }])
    列映射 = {
        "date": "日期",
        "summary": "摘要",
        "mode": "debit_credit",
        "debit": "本位币借方金额",
        "credit": "本位币贷方金额",
        "account": "本方账号",
        "currency": "本位币币种",
        "amount_basis": "本位币",
        "auxiliary_text_columns": ["摘要"],
    }

    标准账 = DataLoader().standardize_data(账, 列映射, "journal")
    金额证据 = 标准账.iloc[0]["amount_evidence"]
    范围证据 = 标准账.iloc[0]["scope_evidence"]

    assert 金额证据["related_amount_values"] == {
        "原币借方金额": 100,
        "原币贷方金额": 0,
        "本位币借方金额": 720,
        "本位币贷方金额": 0,
    }
    assert 范围证据["account_column"] == "本方账号"
    assert 范围证据["account_value"] == "A001"
    assert 范围证据["currency_column"] == "本位币币种"
    assert 范围证据["currency_value"] == "CNY"
    assert 范围证据["related_currency_values"] == {
        "原币币种": "USD",
        "本位币币种": "CNY",
    }


def test_凭证号每行重复而凭证字合并留空时续行仍继承凭证字():
    账 = 原始([400, 600])
    账["日期"] = ["2026-08-10", ""]
    账["凭证字"] = ["记", ""]
    账["凭证号"] = ["001", "001"]

    标准账 = DataLoader().standardize_data(账, 映射(), "journal")

    assert 标准账["voucher_word"].tolist() == ["记", "记"]
    assert 标准账["voucher_no"].tolist() == ["001", "001"]
    assert 标准账.iloc[1]["voucher_evidence"]["inherited"] is True


def test_同一工资期间一侧省略年份仍保留有差额的完整批次():
    银行 = 原始([100, 100, 100, 100], "bank", "8月工资")
    账 = 原始([390], 摘要="2026年8月工资汇总")

    实例 = 核对(银行, 账)

    assert len(实例.selected_candidates) == 1
    候选 = 实例.selected_candidates[0]
    assert 候选.bank_idxs == (0, 1, 2, 3) and 候选.journal_idxs == (0,)
    assert 候选.metrics.total_diff_li == PrecisionEngine.to_integer_li(10)
    assert 候选.processing_status is ProcessingStatus.AUTO_CLASSIFIED
