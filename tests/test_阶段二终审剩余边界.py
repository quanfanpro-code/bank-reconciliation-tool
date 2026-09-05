"""第二阶段终审：用途边界、凭证编号、强证据优先级与高重复性能。"""
from decimal import Decimal

import pandas as pd

import matcher as matcher_module
from data_structures import MatcherConfig, ProcessingStatus
from matcher import Matcher
from precision_engine import PrecisionEngine
from 业务分组 import complete_groups, row_business


def _记录(金额, 摘要, 对方="甲公司", 凭证="", 备注=""):
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


def test_多对多同额但具体用途互不相容时不得仅凭同一主体整组确认():
    匹配器 = _核对(
        [
            _记录(-500, "设备采购款"),
            _记录(-500, "材料采购款"),
        ],
        [
            _记录(-500, "借款归还"),
            _记录(-500, "房屋租金"),
        ],
    )

    assert not any(
        候选.match_type == "closed_candidate_group"
        and 候选.processing_status is ProcessingStatus.GROUP_RECONCILED
        for 候选 in 匹配器.selected_candidates
    )


def test_同月同凭证字但凭证号分隔符不同必须保持两张完整凭证():
    记录集 = [
        _记录(-40, "项目结算", 凭证="001-01"),
        _记录(-60, "项目结算", 凭证="00101"),
    ]
    业务画像 = {
        序号: row_business(行)
        for 序号, 行 in _表格(记录集).iterrows()
    }

    凭证组 = [
        (键, 索引)
        for 键, 索引 in complete_groups(业务画像, 31, "journal")
        if 键[0] == "凭证"
    ]
    assert len(凭证组) == 2
    assert {键[3] for 键, _ in 凭证组} == {"001-01", "00101"}
    assert {索引 for _, 索引 in 凭证组} == {(0,), (1,)}


def test_截断候选必须先保留同主体且同摘要的唯一最佳去向():
    银行记录 = _记录(100, "唯一正确用途", 备注="银行")
    候选正确行 = [
        _记录(100, "唯一正确用途", 备注=f"正确{序号:03d}")
        for 序号 in range(200)
    ]
    正确行 = max(
        候选正确行,
        key=lambda 行: Matcher._row_content_hash(_表格([行]), 0),
    )
    正确哈希 = Matcher._row_content_hash(_表格([正确行]), 0)
    备选行 = [
        _记录(100, f"其他用途{序号:04d}", 备注=f"备选{序号:04d}")
        for 序号 in range(2000)
    ]
    哈希更小的备选 = [
        行 for 行 in 备选行
        if Matcher._row_content_hash(_表格([行]), 0) < 正确哈希
    ][:30]
    assert len(哈希更小的备选) == 30

    匹配器 = _核对(
        [银行记录],
        [*哈希更小的备选, 正确行],
        max_candidates=30,
    )
    正确索引 = 30
    assert any(
        候选.match_type == "exact_1to1"
        and 候选.journal_idxs == (正确索引,)
        for 候选 in 匹配器.candidates
    )
    assert 匹配器.selected_candidates[0].journal_idxs == (正确索引,)


def test_高重复完整运行的歧义优先级计算不随候选数平方增长(monkeypatch):
    原函数 = matcher_module.relationship_priority
    调用数 = 0

    def 计数(候选):
        nonlocal 调用数
        调用数 += 1
        return 原函数(候选)

    monkeypatch.setattr(matcher_module, "relationship_priority", 计数)
    银行 = [
        _记录(100, "同额项目", 备注=f"银行{序号:02d}")
        for 序号 in range(40)
    ]
    日记账 = [
        _记录(100, "同额项目", 备注=f"日记账{序号:02d}")
        for 序号 in range(40)
    ]

    匹配器 = _核对(银行, 日记账, max_candidates=30)

    assert len(匹配器.candidates) >= 1200
    assert 调用数 <= len(匹配器.candidates) * 4
