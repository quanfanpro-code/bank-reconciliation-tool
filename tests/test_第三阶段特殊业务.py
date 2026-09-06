from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter
from 业务事件 import detect_same_side_events


def _df(rows):
    records = []
    for index, (date, amount, fields) in enumerate(rows):
        fields = dict(fields)
        records.append(
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": fields.get("摘要", ""),
                "aux_text_fields": fields,
                "original_idx": index,
                "original_file_row": index + 2,
                "voucher_no": fields.get("凭证号", ""),
                "transaction_id": fields.get("流水号", ""),
            }
        )
    return pd.DataFrame(
        records,
        columns=[
            "date", "amount", "amount_decimal", "summary",
            "aux_text_fields", "original_idx", "original_file_row",
            "voucher_no", "transaction_id",
        ],
    )


def test_付款退回后重付形成完整同侧业务链():
    bank = _df(
        [
            ("2025-12-31", -1000, {"摘要": "甲公司货款", "对方": "甲公司", "业务编号": "P01"}),
            ("2026-01-02", 1000, {"摘要": "甲公司退汇", "对方": "甲公司", "业务编号": "P01"}),
            ("2026-01-03", -1000, {"摘要": "甲公司重新付款", "对方": "甲公司", "业务编号": "P01"}),
        ]
    )

    events, clues = detect_same_side_events(bank, "bank", 31)

    assert len(events) == 1
    assert events[0].event_type == "退汇后重付"
    assert events[0].row_idxs == (0, 1, 2)
    assert events[0].net_amount_li == PrecisionEngine.to_integer_li("-1000")
    assert events[0].is_cross_period is True
    assert clues == []


def test_正负同额但对方不同只形成疑点线索():
    bank = _df(
        [
            ("2026-01-02", -1000, {"摘要": "付款", "对方": "甲公司"}),
            ("2026-01-03", 1000, {"摘要": "退款", "对方": "乙公司"}),
        ]
    )

    events, clues = detect_same_side_events(bank, "bank", 31)

    assert events == []
    assert len(clues) == 1
    assert clues[0].clue_type == "正负同额线索"
    assert "对方不一致" in clues[0].reason


def test_相同流水号形成强重复线索但不删除原行():
    bank = _df(
        [
            ("2026-01-02", -1000, {"摘要": "付款", "对方": "甲公司", "流水号": "TX01"}),
            ("2026-01-02", -1000, {"摘要": "付款", "对方": "甲公司", "流水号": "TX01"}),
        ]
    )

    events, clues = detect_same_side_events(bank, "bank", 31)

    assert events == []
    assert len(clues) == 1
    assert clues[0].clue_type == "强重复线索"
    assert clues[0].row_idxs == (0, 1)
    assert len(bank) == 2


def test_内容相同但流水号不同只形成同额重复疑点():
    bank = _df(
        [
            ("2026-01-02", -1000, {"摘要": "付款", "对方": "甲公司", "流水号": "TX01"}),
            ("2026-01-02", -1000, {"摘要": "付款", "对方": "甲公司", "流水号": "TX02"}),
        ]
    )

    _events, clues = detect_same_side_events(bank, "bank", 31)

    assert len(clues) == 1
    assert clues[0].clue_type == "同额重复疑点"


def test_原付款与明确退款有共同业务证据时形成两段业务链():
    bank = _df(
        [
            ("2026-01-02", -1000, {"摘要": "支付甲公司货款", "对方": "甲公司", "业务编号": "P03"}),
            ("2026-01-05", 1000, {"摘要": "甲公司退款", "对方": "甲公司", "业务编号": "P03"}),
        ]
    )

    events, clues = detect_same_side_events(bank, "bank", 31)

    assert len(events) == 1
    assert events[0].event_type == "付款退回"
    assert events[0].net_amount_li == 0
    assert clues == []


def test_手续费净额有明确费用文字时形成公式关系():
    bank = _df(
        [("2026-01-05", 990, {"摘要": "甲公司回款", "对方": "甲公司", "业务编号": "P02"})]
    )
    journal = _df(
        [
            ("2026-01-05", 1000, {"摘要": "甲公司货款", "对方": "甲公司", "业务编号": "P02"}),
            ("2026-01-05", -10, {"摘要": "银行手续费", "业务编号": "P02"}),
        ]
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()
    result = [item for item in matcher.selected_candidates if item.match_type == "fee_net"]

    assert len(result) == 1
    assert result[0].metrics.total_diff_li == 0
    assert result[0].evidence["relationship_formula"] == "银行实收＝账面应收－手续费"
    assert result[0].evidence["fee_amount_li"] == PrecisionEngine.to_integer_li("10")


def test_同额普通支出不能被当成手续费凑净额():
    bank = _df([("2026-01-05", 990, {"摘要": "甲公司回款", "对方": "甲公司"})])
    journal = _df(
        [
            ("2026-01-05", 1000, {"摘要": "甲公司货款", "对方": "甲公司"}),
            ("2026-01-05", -10, {"摘要": "办公用品"}),
        ]
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()

    assert not any(item.match_type == "fee_net" for item in matcher.candidates)


def test_特殊业务报告包含业务链净额公式截止性和重复线索():
    bank = _df(
        [
            ("2025-12-31", -1000, {"摘要": "甲公司货款", "对方": "甲公司", "业务编号": "P01"}),
            ("2026-01-02", 1000, {"摘要": "甲公司退汇", "对方": "甲公司", "业务编号": "P01"}),
            ("2026-01-03", -1000, {"摘要": "甲公司重付", "对方": "甲公司", "业务编号": "P01"}),
        ]
    )
    journal = _df([])
    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()

    tables = Reporter(matcher).build_report_tables(MatcherConfig())

    assert {"退款冲销重付", "手续费及净额", "截止性差异", "重复线索"}.issubset(tables)
    assert len(tables["退款冲销重付"]) == 3
    assert "关系公式" in tables["手续费及净额"].columns
    assert len(tables["截止性差异"]) == 3
