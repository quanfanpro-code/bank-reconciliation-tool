from decimal import Decimal
from pathlib import Path

import pytest

from gui import (
    ColumnMappingFrame,
    ReconciliationApp,
    _auto_mapping_for_columns,
    auto_select_auxiliary_columns,
    build_llm_config,
    build_matcher_config,
    default_gui_values,
)


def test_GUI使用三个完整名称和已确认默认值():
    values = default_gui_values()
    config = build_matcher_config(values)

    assert values["performance_materiality"] == "100000"
    assert values["clearly_trivial_threshold"] == "5000"
    assert values["auto_confirm_score"] == "70"
    assert config.performance_materiality == Decimal("100000.00")
    assert config.clearly_trivial_threshold == Decimal("5000.00")
    assert config.auto_confirm_score == 70


def test_辅助字段自动选择但账号卡号凭证号不默认选择():
    columns = [
        "摘要",
        "业务说明",
        "交易用途",
        "对方户名",
        "附言",
        "备注",
        "账号",
        "卡号",
        "凭证号",
    ]

    selected = auto_select_auxiliary_columns(columns)

    assert selected == [
        "摘要",
        "业务说明",
        "交易用途",
        "对方户名",
        "附言",
        "备注",
    ]


def test_自动列映射能识别单列金额模式而不误认借贷分列():
    with_direction = _auto_mapping_for_columns(
        ["日期", "金额", "借贷标志", "摘要"],
        is_bank=True,
    )
    signed = _auto_mapping_for_columns(
        ["日期", "发生金额", "摘要"],
        is_bank=True,
    )

    assert with_direction["mode"] == "single_amount_with_direction"
    assert with_direction["amount"] == "金额"
    assert with_direction["direction"] == "借贷标志"
    assert signed["mode"] == "signed_amount"


def test_用户明确不选辅助字段时重新打开仍保持为空():
    try:
        app = ReconciliationApp()
    except Exception as exc:
        if "display" in str(exc).lower() or "tcl" in str(exc).lower():
            pytest.skip(f"当前环境无法显示图形窗口：{exc}")
        raise
    try:
        app.withdraw()
        frame = ColumnMappingFrame(app, "测试")
        frame.set_columns(["日期", "借方", "贷方", "摘要", "业务说明"])
        frame.set_mapping(
            {
                "mode": "debit_credit",
                "date": "日期",
                "debit": "借方",
                "credit": "贷方",
                "auxiliary_text_columns": [],
            }
        )

        assert frame.get_mapping()["auxiliary_text_columns"] == []
    finally:
        app.destroy()


def test_大模型关闭时地址模型和密钥可以为空():
    config = build_llm_config(
        {
            "enabled": False,
            "mode": "online",
            "protocol": "auto",
            "base_url": "",
            "model": "",
            "api_key": "",
            "timeout_seconds": "",
            "candidate_limit": "",
            "local_fields": "",
        }
    )

    assert config.enabled is False
    assert config.api_key == ""


def test_LM_Studio默认地址和用户选择字段正确():
    config = build_llm_config(
        {
            "enabled": True,
            "mode": "local",
            "protocol": "auto",
            "base_url": "",
            "model": "local-model",
            "api_key": "",
            "timeout_seconds": "30",
            "candidate_limit": "5",
            "local_fields": "摘要, 对方户名, 账号",
        }
    )

    assert config.base_url == "http://localhost:1234/v1"
    assert config.local_fields == ("摘要", "对方户名", "账号")


def test_本地大模型字段收到异常格式时安全按空列表处理():
    config = build_llm_config(
        {
            "enabled": True,
            "mode": "local",
            "protocol": "auto",
            "base_url": "",
            "model": "local-model",
            "api_key": "",
            "timeout_seconds": "30",
            "candidate_limit": "5",
            "local_fields": object(),
        }
    )

    assert config.local_fields == ()


def test_启用在线大模型时校验地址模型超时和候选数():
    base = {
        "enabled": True,
        "mode": "online",
        "protocol": "auto",
        "base_url": "https://example.com/v1",
        "model": "model",
        "api_key": "session-only",
        "timeout_seconds": "30",
        "candidate_limit": "5",
        "local_fields": "",
    }

    assert build_llm_config(base).model == "model"
    with pytest.raises(ValueError, match="服务地址"):
        build_llm_config({**base, "base_url": "not-a-url"})
    with pytest.raises(ValueError, match="模型"):
        build_llm_config({**base, "model": ""})
    with pytest.raises(ValueError, match="超时"):
        build_llm_config({**base, "timeout_seconds": "0"})
    with pytest.raises(ValueError, match="候选"):
        build_llm_config({**base, "candidate_limit": "0"})


def test_API密钥不会进入匹配器参数():
    values = default_gui_values()
    values["api_key"] = "绝不能进入匹配参数"

    config = build_matcher_config(values)

    assert not hasattr(config, "api_key")


def test_主窗口保留完整策略名称和独立配置入口():
    source = Path("gui.py").read_text(encoding="utf-8")

    assert "实际执行重要性水平" in source
    assert "明显微小错报临界值" in source
    assert "配置列映射与辅助字段" in source
    assert "高级匹配参数" in source
    assert "配置与测试" in source
    assert "本机生成的敏感字段相同/不同结论" in source
    assert "不发送账号、卡号等原值" in source


def test_主窗口尺寸和日志伸缩规则():
    source = Path("gui.py").read_text(encoding="utf-8")

    assert 'self.geometry("1200x850")' in source
    assert "self.minsize(960, 700)" in source
    assert "self.grid_rowconfigure(3, weight=1)" in source


def test_主窗口可以完成最小启动():
    try:
        app = ReconciliationApp()
    except Exception as exc:
        if "display" in str(exc).lower() or "tcl" in str(exc).lower():
            pytest.skip(f"当前环境无法显示图形窗口：{exc}")
        raise
    try:
        app.update()
        assert app.winfo_width() >= 960
        assert app.winfo_height() >= 700
        assert app.txt_log.winfo_exists()
    finally:
        app.destroy()


def test_图形界面复用无界面的完整核对流程():
    source = Path("gui.py").read_text(encoding="utf-8")

    assert "from application import run_reconciliation" in source
    assert "output_path = run_reconciliation(" in source
