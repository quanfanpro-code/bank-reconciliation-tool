# -*- coding: utf-8 -*-
"""余额断档局部降级：窗口标注、局部化判定、回退。"""
from decimal import Decimal

import pandas as pd

from balance import build_overall_controls, check_balance_continuity


def _流水(行):
    frame = pd.DataFrame(行, columns=["date", "amount", "balance", "original_file_row"])
    frame["date"] = pd.to_datetime(frame["date"])
    return frame


def test_断档异常携带窗口日期():
    frame = _流水([
        ("2026-01-05", Decimal("5000"), Decimal("5000"), 2),
        # 01-06 开盘倒推 4000，与上日末 5000 差 1000（真实缺行）
        ("2026-01-06", Decimal("100"), Decimal("4100"), 3),
        ("2026-01-07", Decimal("50"), Decimal("4150"), 4),
    ])
    异常 = check_balance_continuity(frame, source="银行流水")
    assert len(异常) == 1
    assert str(异常[0]["窗口止"])[:10] == "2026-01-06"
    assert str(异常[0]["窗口起"])[:10] == "2026-01-05"


def test_断档可定位且不过半时不触发整体降级():
    行 = [
        ("2026-01-05", Decimal("5000"), Decimal("5000"), 2),
        ("2026-01-06", Decimal("100"), Decimal("4100"), 3),
        ("2026-01-07", Decimal("50"), Decimal("4150"), 4),
        ("2026-01-08", Decimal("60"), Decimal("4210"), 5),
        ("2026-01-09", Decimal("70"), Decimal("4280"), 6),
    ]
    银行 = _流水(行)
    账 = 银行.drop(columns=["balance"]).copy()
    控制 = build_overall_controls(银行, 账)
    assert 控制.scope_limited is False
    assert len(控制.affected_windows) == 1
    起, 止 = 控制.affected_windows[0]
    assert str(起)[:10] == "2026-01-05"
    assert str(止)[:10] == "2026-01-06"
    # 断档原因仍披露在 reasons 里
    assert any("连续性异常" in 原因 for 原因 in 控制.reasons)


def test_断档窗口覆盖过半时回退整体降级():
    银行 = _流水([
        ("2026-01-05", Decimal("5000"), Decimal("5000"), 2),
        ("2026-01-06", Decimal("100"), Decimal("4100"), 3),
        ("2026-01-07", Decimal("50"), Decimal("4150"), 4),
    ])
    账 = 银行.drop(columns=["balance"]).copy()
    控制 = build_overall_controls(银行, 账)
    assert 控制.scope_limited is True  # 3 个交易日中窗口覆盖 2 日，过半回退


from application import run_reconciliation
from data_structures import MatcherConfig, ProcessingStatus

余额映射 = {
    "date": "日期", "amount": "金额", "balance": "余额", "summary": "摘要",
    "mode": "signed_amount", "account": "本方账号", "currency": "币种",
    "amount_basis": "本位币", "auxiliary_text_columns": ["摘要", "业务编号"],
}


def _余额行(日期, 金额, 摘要, 编号, 余额):
    return {"日期": 日期, "金额": 金额, "摘要": 摘要, "业务编号": 编号,
            "本方账号": "A001", "币种": "CNY", "余额": 余额}


def _写带余额表(tmp_path, 银行行, 账行):
    银行路径 = tmp_path / "银行.csv"
    账路径 = tmp_path / "账.csv"
    pd.DataFrame(银行行).to_csv(银行路径, index=False, encoding="utf-8-sig")
    pd.DataFrame(账行).to_csv(账路径, index=False, encoding="utf-8-sig")
    输出 = tmp_path / "核对结果.xlsx"
    捕获 = []
    run_reconciliation(
        str(银行路径), str(账路径), 余额映射, 余额映射, MatcherConfig(),
        bank_skiprows=0, journal_skiprows=0, bank_header_rows=1, journal_header_rows=1,
        output_path=str(输出), matcher_ready=捕获.append,
    )
    return 输出, 捕获[0]


def test_断档窗口内降级窗口外正常确认(tmp_path):
    银行 = [
        _余额行("2026-08-10", 100, "甲公司回款", "P01", 1100),
        # 断档：08-11 开盘倒推 700，与 08-10 收盘 1100 差 400
        _余额行("2026-08-11", 200, "乙公司回款", "P02", 900),
        _余额行("2026-08-12", 300, "丙公司回款", "P03", 1200),
        _余额行("2026-08-13", -50, "付丁公司款", "P04", 1150),
        _余额行("2026-08-14", 70, "戊公司回款", "P05", 1220),
    ]
    账 = [dict(r, 余额="") for r in 银行]
    _, 匹配器 = _写带余额表(tmp_path, 银行, 账)

    assert 匹配器.overall_control.scope_limited is False
    assert len(匹配器.overall_control.affected_windows) == 1
    降级 = [c for c in 匹配器.selected_candidates if c.processing_status is ProcessingStatus.FLAGGED]
    确认 = [c for c in 匹配器.selected_candidates if c.processing_status is ProcessingStatus.AUTO_CONFIRMED]
    assert len(降级) == 2, f"窗口内P01/P02应降级，实际{len(降级)}笔"
    assert len(确认) == 3, f"窗口外P03~P05应自动确认，实际{len(确认)}笔"
    assert any(
        "断档" in str(理由)
        for c in 降级 for 理由 in c.evidence.get("overall_control_reasons", ())
    )


def test_跨窗口匹配一侧落入即降级(tmp_path):
    银行 = [
        _余额行("2026-08-10", 100, "甲公司回款", "P01", 1100),
        # 断档：08-11 开盘倒推 400，与 08-10 收盘 1100 差 700
        _余额行("2026-08-11", 500, "乙公司回款", "P02", 900),
        _余额行("2026-08-13", 300, "丙公司回款", "P03", 1200),
    ]
    账 = [
        _余额行("2026-08-12", 500, "乙公司回款", "P02", ""),
        _余额行("2026-08-13", 300, "丙公司回款", "P03", ""),
    ]
    _, 匹配器 = _写带余额表(tmp_path, 银行, 账)

    降级 = [c for c in 匹配器.selected_candidates if c.processing_status is ProcessingStatus.FLAGGED]
    确认 = [c for c in 匹配器.selected_candidates if c.processing_status is ProcessingStatus.AUTO_CONFIRMED]
    # P02 跨窗口（银行侧 08-11 在窗口内）降级；P03 两侧都在窗口外正常确认
    assert len(降级) == 1, f"P02应降级，实际{len(降级)}笔"
    assert len(确认) == 1, f"P03应自动确认，实际{len(确认)}笔"


def test_系统结论披露断档位置而非全部降级(tmp_path):
    银行 = [
        _余额行("2026-08-10", 100, "甲公司回款", "P01", 1100),
        _余额行("2026-08-11", 200, "乙公司回款", "P02", 900),
        _余额行("2026-08-12", 300, "丙公司回款", "P03", 1200),
        _余额行("2026-08-13", -50, "付丁公司款", "P04", 1150),
        _余额行("2026-08-14", 70, "戊公司回款", "P05", 1220),
    ]
    账 = [dict(r, 余额="") for r in 银行]
    输出, _ = _写带余额表(tmp_path, 银行, 账)

    报表 = pd.read_excel(输出, sheet_name=None)
    结论表 = 报表["核对结论"]
    系统结论 = str(结论表.loc[结论表["项目"] == "系统结论", "数值"].iloc[0])
    assert "断档" in 系统结论 and "2026-08-11" in 系统结论
    assert "全部" not in 系统结论 and "总体余额控制异常" not in 系统结论
    检查项 = 报表["输入检查"].set_index("检查项目")
    assert "断档窗口" in str(检查项.at["总体余额控制", "说明"])
