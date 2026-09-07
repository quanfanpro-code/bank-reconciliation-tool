"""独立复核：金额、日期和外部回答的异常不能冒充正常数据。"""

from decimal import Decimal

import pandas as pd
import pytest

from data_loader import DataLoader, ParseErrorCollector
from llm_assistant import LLMConfig, LLMAssistant
from tests.test_llm_assistant import _sample_request
from utils import parse_date
from validate import validate_config_params


@pytest.mark.parametrize("模式", ["signed_amount", "single_amount_with_direction", "debit_credit"])
@pytest.mark.parametrize("异常值", [float("inf"), Decimal("1e100"), "NaN123"])
def test_坏金额逐行留证且正常行继续处理(模式, 异常值):
    原表 = pd.DataFrame({"日期": ["2026-01-02", "2026-01-03"],
                        "金额": [异常值, 100], "方向": ["收入", "收入"], "支出": [0, 0]})
    映射 = {"date": "日期", "mode": 模式, "amount": "金额", "direction": "方向",
            "credit": "金额", "debit": "支出"}
    收集器 = ParseErrorCollector()
    结果 = DataLoader(error_collector=收集器).standardize_data(原表, 映射, "bank")
    assert list(结果["original_file_row"]) == [3]
    assert list(结果["amount"]) == [Decimal("100.00")]
    assert any(项["type"] == "金额解析失败" and 项["row"] == 2
               for 项 in 收集器.get_all_errors())


@pytest.mark.parametrize("值", [20260102.0, "20260102.0"])
def test_数值八位日期不因浮点显示丢失(值):
    assert parse_date(值) == pd.Timestamp("2026-01-02")


@pytest.mark.parametrize("参数,值", [("performance_materiality", "nan"),
                                    ("clearly_trivial_threshold", "inf"),
                                    ("memory_limit", "inf"), ("random_seed", "--1"),
                                    ("auto_confirm_score", "²")])
def test_非法界面参数清楚拒绝且不崩溃(参数, 值):
    有效, 提示 = validate_config_params("31", "31", "30", "3", **{参数: 值})
    assert not 有效 and 提示


@pytest.mark.parametrize("协议,回答", [("responses", {"output": None}),
                                     ("responses", {"output": [{"content": None}]}),
                                     ("chat_completions", {"choices": {"bad": 1}})])
def test_服务端异常结构仍回退本地规则(monkeypatch, 协议, 回答):
    助手 = LLMAssistant(LLMConfig(enabled=True, protocol=协议,
                               base_url="http://127.0.0.1:1/v1", model="本地测试"))
    monkeypatch.setattr(助手, "_call_protocol", lambda *args: 回答)
    结果 = 助手.evaluate_candidates(_sample_request())
    assert 结果.fallback_used and 结果.error
    assert 结果.selected_candidate_id == ""


@pytest.mark.parametrize("扩展名", [".csv", ".xlsx"])
def test_输入编号中的NA与NULL是原文而不是空值(tmp_path, 扩展名):
    路径 = tmp_path / ("原文" + 扩展名)
    表 = pd.DataFrame({"日期": ["2026-01-02", "2026-01-03"],
                       "金额": [100, 200], "业务编号": ["NA", "NULL"]})
    if 扩展名 == ".csv":
        表.to_csv(路径, index=False, encoding="utf-8-sig")
    else:
        表.to_excel(路径, index=False)
    结果 = DataLoader().load_file(str(路径))
    assert list(结果["业务编号"]) == ["NA", "NULL"]
