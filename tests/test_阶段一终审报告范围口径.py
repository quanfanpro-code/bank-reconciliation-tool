"""阶段一终审：辅助文字提示不能把已闭合的核对范围误写为受限。"""

from decimal import Decimal

import pandas as pd

from data_structures import MatcherConfig, OverallControlResult
from input_precheck import InputPrecheckReport, PrecheckItem
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter


def test_只有辅助文字完整性提示时首页仍应显示范围可用():
    表 = pd.DataFrame([
        {
            "date": pd.Timestamp("2026-08-10"),
            "amount": Decimal("100"),
            "amount_decimal": PrecisionEngine.to_integer_li(100),
            "summary": "项目回款",
            "aux_text_fields": {"摘要": "项目回款"},
            "voucher_word": "",
            "voucher_no": "",
            "original_idx": 1,
            "original_file_row": 2,
        }
    ])
    总体控制 = OverallControlResult(
        period_status="通过",
        amount_status="通过",
        scope_limited=False,
    )
    预检 = InputPrecheckReport(
        items=(
            PrecheckItem(
                "辅助文字完整性",
                "摘要100%；批次号0%",
                "摘要100%；批次号0%",
                "至少一侧偏低",
                "疑点",
                "可用文字较少，匹配证据可能不足。",
            ),
        ),
        overall_control=总体控制,
    )
    匹配器 = Matcher(
        表,
        表.copy(),
        MatcherConfig(),
        overall_control=总体控制,
        logger=lambda _: None,
    )
    匹配器.run()

    核对结论 = Reporter(
        匹配器,
        precheck_report=预检,
    ).build_report_tables(匹配器.config)["核对结论"]
    首页 = dict(zip(核对结论["项目"], 核对结论["数值"]))

    assert 首页["核对范围"] == "范围可用"
    assert 首页["范围说明"] == "未发现范围疑点"
    assert "核对已自动完成" in str(首页["系统结论"])
