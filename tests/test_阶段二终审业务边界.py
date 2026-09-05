"""第二阶段终审发现的批次、凭证和换序边界。"""
from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter


def _记录(金额, 摘要="支付货款", 对方="甲公司", 凭证="", **辅助字段):
    文字字段 = {"摘要": 摘要}
    if 对方:
        文字字段["对方户名"] = 对方
    文字字段.update(辅助字段)
    return {
        "date": pd.Timestamp("2026-08-10"),
        "amount": Decimal(str(金额)),
        "amount_decimal": PrecisionEngine.to_integer_li(金额),
        "summary": 摘要,
        "aux_text_fields": 文字字段,
        "voucher_word": "记" if 凭证 else "",
        "voucher_no": 凭证,
    }


def _表格(记录集):
    return pd.DataFrame(
        [
            dict(行, original_idx=序号 + 1, original_file_row=序号 + 2)
            for 序号, 行 in enumerate(记录集)
        ]
    )


def _核对(银行记录, 日记账记录, **配置):
    匹配器 = Matcher(
        _表格(银行记录),
        _表格(日记账记录),
        MatcherConfig(**配置),
        logger=lambda _: None,
    )
    匹配器.run()
    return 匹配器


def test_明确批次号有差额时必须生成批次核查线索并进入两张复核表():
    匹配器 = _核对(
        [
            _记录(-40, 批次号="B-001"),
            _记录(-50, 批次号="B-001"),
        ],
        [_记录(-100, 批次号="B-001")],
    )

    候选 = next(
        候选
        for 候选 in 匹配器.selected_candidates
        if 候选.bank_idxs == (0, 1) and 候选.journal_idxs == (0,)
    )
    assert 候选.evidence.get("batch_difference") is True
    assert 候选.evidence.get("batch_bank_count") == 2
    assert 候选.evidence.get("batch_journal_count") == 1
    assert 候选.evidence.get("batch_difference_li") == PrecisionEngine.to_integer_li(10)
    assert 候选.evidence.get("batch_possible_missing_side")
    assert 候选.evidence.get("batch_review_hint")

    报表 = Reporter(匹配器).build_report_tables(匹配器.config)
    for 表名 in ("整组勾稽", "疑点事项"):
        行 = 报表[表名].loc[报表[表名]["匹配ID"] == 候选.final_match_id]
        assert len(行) == 1
        assert 行.iloc[0]["批次核查线索"]


def test_强业务证据使大额差异仍保留完整凭证关系并列明组成():
    匹配器 = _核对(
        [_记录(-200000, 摘要="项目结算", 对方="甲公司")],
        [
            _记录(-25000, 摘要="项目结算", 对方="甲公司", 凭证="001"),
            _记录(-25000, 摘要="项目结算", 对方="甲公司", 凭证="001"),
        ],
        performance_materiality=Decimal("100000"),
    )

    候选 = next(
        候选
        for 候选 in 匹配器.selected_candidates
        if 候选.bank_idxs == (0,) and 候选.journal_idxs == (0, 1)
    )
    assert 候选.metrics.total_diff_li == PrecisionEngine.to_integer_li(150000)
    assert 候选.evidence.get("atomic_voucher_group") is True
    assert 候选.evidence.get("batch_difference") is True
    assert 候选.evidence.get("batch_review_hint")

    报表 = Reporter(匹配器).build_report_tables(匹配器.config)
    组成 = 报表["匹配组成"].loc[报表["匹配组成"]["匹配ID"] == 候选.final_match_id]
    assert len(组成) == 3
    assert set(组成["来源"]) == {"银行流水", "日记账"}


def test_精确同额候选超过上限时换序不改变候选和最终关系编号():
    银行 = [_记录(100, 摘要="同额测试", 对方="")]
    日记账 = [
        _记录(100, 摘要="同额测试", 对方="", 备注=f"T{序号:02d}")
        for 序号 in range(31)
    ]

    def 结果(记录集):
        匹配器 = _核对(银行, 记录集, max_candidates=30)
        稳定键 = {
            候选.stable_key
            for 候选 in 匹配器.candidates
            if 候选.match_type == "exact_1to1"
        }
        最终 = 匹配器.selected_candidates[0]
        return 稳定键, 最终.composition_key, 最终.final_match_id

    内容稳定最小 = min(
        日记账,
        key=lambda 行: Matcher._row_content_hash(_表格([行]), 0),
    )
    其余 = [行 for 行 in 日记账 if 行 is not 内容稳定最小]
    尾部顺序 = [*其余, 内容稳定最小]
    首部顺序 = [内容稳定最小, *其余]

    assert 结果(尾部顺序) == 结果(首部顺序)
