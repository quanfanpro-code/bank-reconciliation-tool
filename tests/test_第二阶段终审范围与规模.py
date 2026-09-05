"""第二阶段终审：范围警告、零金额、凭证证据和密集多解必须保守且可扩展。"""
from decimal import Decimal

import pandas as pd
import pytest

from application import run_reconciliation
from data_structures import MatcherConfig, ProcessingStatus, RiskLevel
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter


def _记录(金额, 摘要="项目收款", 对方="甲公司", 凭证="", 备注=""):
    辅助字段 = {"摘要": 摘要}
    if 对方:
        辅助字段["对方户名"] = 对方
    if 备注:
        辅助字段["备注"] = 备注
    return {
        "date": pd.Timestamp("2026-08-10"),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 辅助字段,
        "voucher_word": "记" if 凭证 else "",
        "voucher_no": 凭证,
    }


def _表格(记录集):
    return pd.DataFrame([
        dict(行, original_idx=序号 + 1, original_file_row=序号 + 2)
        for 序号, 行 in enumerate(记录集)
    ])


def _核对(银行, 日记账, **配置):
    匹配器 = Matcher(
        _表格(银行),
        _表格(日记账),
        MatcherConfig(**配置),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def test_完整凭证缺少跨侧业务证据时不得自动确认():
    匹配器 = _核对(
        [_记录(-100000, "设备采购款", 对方="")],
        [
            _记录(-40000, "房屋租金", 对方="", 凭证="001"),
            _记录(-60000, "房屋租金", 对方="", 凭证="001"),
        ],
    )

    凭证候选 = next(
        候选 for 候选 in 匹配器.candidates
        if 候选.bank_idxs == (0,)
        and 候选.journal_idxs == (0, 1)
        and 候选.evidence.get("atomic_voucher_group")
    )
    assert 凭证候选.evidence.get("resolves_full_group") is False
    assert 凭证候选.evidence.get("complete_business_group") is False
    assert 凭证候选.processing_status is ProcessingStatus.FLAGGED
    assert not any(
        候选.processing_status in {
            ProcessingStatus.AUTO_CONFIRMED,
            ProcessingStatus.GROUP_RECONCILED,
        }
        for 候选 in 匹配器.selected_candidates
    )


def test_默认不得建立纯零金额关系但明确允许时可以建立():
    默认 = _核对([_记录(0)], [_记录(0)])
    明确允许 = _核对([_记录(0)], [_记录(0)], allow_zero_match=True)

    assert not 默认.candidates
    assert not 默认.selected_candidates
    assert any(
        候选.bank_idxs == (0,) and 候选.journal_idxs == (0,)
        for 候选 in 明确允许.candidates
    )


def test_双方均有非零收支但净额为零的组不受纯零金额开关误伤():
    匹配器 = Matcher(
        _表格([_记录(100), _记录(-100)]),
        _表格([_记录(60), _记录(-60)]),
        MatcherConfig(allow_zero_match=False),
        logger=lambda _: None,
    )

    候选 = 匹配器._add_candidate([0, 1], [0, 1], "test_group", "测试")
    assert 候选 is not None


@pytest.mark.parametrize("场景", ["账户部分缺失", "币种部分缺失", "金额口径未注明"])
def test_范围身份未完全验证必须贯穿到关系状态和正式报告(tmp_path, 场景):
    银行 = pd.DataFrame({
        "日期": ["2026-08-10", "2026-08-11"],
        "金额": [100, 200],
        "摘要": ["项目甲收款", "项目乙收款"],
        "账号": ["A001", "A001"],
        "币种": ["CNY", "CNY"],
    })
    日记账 = 银行.copy()
    if 场景 == "账户部分缺失":
        银行.loc[1, "账号"] = ""
    elif 场景 == "币种部分缺失":
        银行.loc[1, "币种"] = ""

    银行路径 = tmp_path / f"{场景}_银行.xlsx"
    账路径 = tmp_path / f"{场景}_序时账.xlsx"
    输出路径 = tmp_path / f"{场景}_核对报告.xlsx"
    银行.to_excel(银行路径, index=False)
    日记账.to_excel(账路径, index=False)
    公共映射 = {
        "date": "日期",
        "amount": "金额",
        "summary": "摘要",
        "account": "账号",
        "currency": "币种",
        "mode": "signed_amount",
    }
    银行映射 = dict(公共映射)
    账映射 = dict(公共映射)
    if 场景 != "金额口径未注明":
        银行映射["amount_basis"] = "本位币"
        账映射["amount_basis"] = "本位币"

    捕获 = []
    run_reconciliation(
        str(银行路径),
        str(账路径),
        银行映射,
        账映射,
        MatcherConfig(),
        bank_skiprows=0,
        journal_skiprows=0,
        bank_header_rows=1,
        journal_header_rows=1,
        output_path=输出路径,
        matcher_ready=捕获.append,
    )

    匹配器 = 捕获[0]
    assert 匹配器.overall_control.scope_limited is True
    assert 匹配器.overall_scope_limited is True
    assert 匹配器.balance_integrity_limited is False
    assert any(场景[:2] in 理由 or "金额口径" in 理由 for 理由 in 匹配器.overall_control.reasons)
    assert 匹配器.selected_candidates
    assert all(候选.processing_status is ProcessingStatus.FLAGGED for 候选 in 匹配器.selected_candidates)
    assert all(候选.risk_level is RiskLevel.UNKNOWN for 候选 in 匹配器.selected_candidates)

    报表 = pd.read_excel(输出路径, sheet_name=None)
    系统结论 = str(
        报表["核对结论"].loc[
            报表["核对结论"]["项目"] == "系统结论", "数值"
        ].iloc[0]
    )
    assert "总体资料" in 系统结论 and "尚未闭合" in 系统结论
    assert "总体余额控制异常" not in 系统结论
    检查项 = 报表["输入检查"].set_index("检查项目")
    项目名 = {
        "账户部分缺失": "核对账户",
        "币种部分缺失": "核对币种",
        "金额口径未注明": "金额口径",
    }[场景]
    assert 检查项.at[项目名, "状态"] == "疑点"
    assert 检查项.at["总体余额控制", "状态"] == "疑点"
    assert not 报表["疑点事项"].empty
    assert set(报表["疑点事项"]["系统结论"]) == {"疑点事项"}
    assert set(报表["疑点事项"]["风险等级"]) == {"范围未知"}


def test_密集同额多解只保留固定数量引用并披露竞争组总数():
    银行 = [_记录(100, 备注=f"银行{序号:03d}") for 序号 in range(40)]
    日记账 = [_记录(100, 备注=f"日记账{序号:03d}") for 序号 in range(40)]
    匹配器 = _核对(银行, 日记账, max_candidates=30)

    assert Matcher.MAX_ALTERNATIVE_REFERENCES == 10
    总引用数 = sum(
        len(候选.evidence.get("alternative_candidate_ids", ()))
        for 候选 in 匹配器.candidates
    )
    assert 总引用数 <= len(匹配器.candidates) * Matcher.MAX_ALTERNATIVE_REFERENCES
    截断候选 = [
        候选 for 候选 in 匹配器.candidates
        if 候选.evidence.get("alternative_candidates_truncated")
    ]
    assert 截断候选
    assert all(
        候选.evidence["alternative_candidate_count"]
        > len(候选.evidence["alternative_candidate_ids"])
        for 候选 in 截断候选
    )
    已选截断 = 截断候选[0]
    说明 = Reporter(匹配器)._alternative_composition(
        已选截断,
        {候选.candidate_id: 候选 for 候选 in 匹配器.candidates},
    )
    assert str(已选截断.evidence["alternative_candidate_count"]) in 说明
    assert "仅列" in 说明
