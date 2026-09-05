"""真实项目运行必须持续显示阶段、数量、耗时和整体进度。"""

from decimal import Decimal

import pandas as pd

from application import _map_matcher_progress
from data_structures import MatcherConfig
from matcher import Matcher
from precision_engine import PrecisionEngine
from reporter import Reporter


def _frame(source: str) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "date": pd.Timestamp("2026-01-05"),
                "amount": Decimal("100"),
                "amount_decimal": PrecisionEngine.to_integer_li(100),
                "summary": "销售回款",
                "aux_text_fields": {
                    "摘要": "销售回款",
                    "对方户名": "甲公司",
                    "业务编号": "P001",
                },
                "voucher_no": "记-001" if source == "journal" else "",
                "source": source,
                "original_idx": 1,
                "original_file_row": 2,
            }
        ]
    )


def _many_frame(source: str, amounts: list[int]) -> pd.DataFrame:
    rows = []
    for index, amount in enumerate(amounts, 1):
        rows.append(
            {
                "date": pd.Timestamp("2026-01-05"),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": "",
                "aux_text_fields": {},
                "voucher_no": "",
                "source": source,
                "original_idx": index,
                "original_file_row": index + 1,
            }
        )
    return pd.DataFrame(rows)


def test_匹配阶段输出完成数量耗时且整体进度不倒退():
    logs: list[str] = []
    progress: list[float] = []
    matcher = Matcher(
        _frame("bank"),
        _frame("journal"),
        MatcherConfig(),
        logger=logs.append,
        progress_callback=progress.append,
    )

    matcher.run()

    expected_stages = (
        "业务完整组匹配",
        "白名单规则匹配",
        "精确匹配",
        "日期容差匹配",
        "批量聚合匹配",
        "连续摘要整组匹配",
        "智能组合匹配",
        "日总额匹配",
        "月度总额匹配",
        "跨月多对多匹配",
    )
    for stage in expected_stages:
        completion = next(line for line in logs if f"完成: {stage}" in line)
        assert "新增候选" in completion
        assert "累计候选" in completion
        assert "耗时" in completion
    assert any("候选选择完成" in line and "选中关系" in line for line in logs)
    assert progress == sorted(progress)
    assert all(value < 100 for value in progress[:-1])
    assert progress[-1] == 100


def test_组合搜索按百分点节流界面回调而不是每个来源任务都排队():
    progress: list[float] = []
    matcher = Matcher(
        _many_frame("bank", [-3] * 40),
        _many_frame("journal", [-1, -2] * 20),
        MatcherConfig(
            combination_global_time_limit_seconds=3,
            combination_task_timeout_seconds=0.05,
            combination_node_limit_per_source=100,
            max_dfs_depth=3,
        ),
        logger=lambda _message: None,
        progress_callback=progress.append,
    )

    matcher.match_dfs_combinations()

    assert progress == sorted(progress)
    assert progress[-1] == 63
    assert len(progress) <= 12


def test_生产入口将匹配器进度线性放进百分之三十至七十五区间():
    assert _map_matcher_progress(0) == 0.3
    assert _map_matcher_progress(50) == 0.525
    assert _map_matcher_progress(100) == 0.75


def test_报告生成实时输出构造写入排版保存节点(tmp_path):
    matcher = Matcher(
        _frame("bank"),
        _frame("journal"),
        MatcherConfig(),
        logger=lambda _message: None,
    )
    matcher.run()
    logs: list[str] = []
    output = tmp_path / "核对报告.xlsx"

    Reporter(matcher, logger=logs.append).generate_report(str(output))

    assert output.is_file()
    assert any("开始构造报告表格" in line for line in logs)
    assert any("报告表格构造完成" in line and "工作表" in line for line in logs)
    assert any("工作簿初次写入完成" in line for line in logs)
    assert any("报告排版保存完成" in line and "耗时" in line for line in logs)
