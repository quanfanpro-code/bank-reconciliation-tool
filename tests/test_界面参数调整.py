"""只检查界面控件和参数传递，不运行核对。"""
from types import SimpleNamespace

import customtkinter as ctk
import gui
from gui import AdvancedConfigDialog, ReconciliationApp, build_matcher_config, default_gui_values


def _控件(root):
    for child in root.winfo_children():
        yield child
        yield from _控件(child)


def test_日期参数在审计策略中且分数门槛在高级设置中正确保存(monkeypatch):
    app = ReconciliationApp()
    app.withdraw()
    dialog = None
    try:
        widgets = list(_控件(app))
        title = next(w for w in widgets if isinstance(w, ctk.CTkLabel) and w.cget("text") == "审计策略")
        card = title.master
        entries = [w for w in _控件(card) if isinstance(w, ctk.CTkEntry)]
        for variable in (app.tolerance_days_var, app.dfs_window_var):
            assert any(str(w.cget("textvariable")) == str(variable) for w in entries)
        assert app.auto_confirm_score_var.get() == "70"
        assert not any(
            isinstance(w, ctk.CTkEntry)
            and str(w.cget("textvariable")) == str(app.auto_confirm_score_var)
            for w in widgets
        )
        app.tolerance_days_var.set("5")
        app.dfs_window_var.set("12")
        before = app._current_advanced_values()
        dialog = AdvancedConfigDialog(app, before)
        dialog.withdraw()
        labels = [str(w.cget("text")) for w in _控件(dialog) if isinstance(w, (ctk.CTkLabel, ctk.CTkCheckBox))]
        assert "自动确认分数门槛" in labels
        assert "自动确认最低综合可信度" not in labels
        assert any("默认70" in s and "冲突" in s and "准确率" in s for s in labels)
        assert dialog.variables["auto_confirm_score"].get() == "70"
        assert "日期读取格式" in labels and "日期显示格式" not in labels
        assert not any(s in labels for s in ("日期容差（天）", "组合窗口（天）", "文字相似度阈值", "高相似度阈值"))
        assert any("固定数字" in s for s in labels)
        assert any("收入" in s or "正数" in s for s in labels)
        assert any("auto" in s and "读取" in s for s in labels)
        dialog.variables["auto_confirm_score"].set("80")
        dialog._save()
        expected = {**before, "auto_confirm_score": "80"}
        assert dialog.result == expected
        assert build_matcher_config(dialog.result) == build_matcher_config(expected)
        monkeypatch.setattr(gui, "AdvancedConfigDialog", lambda parent, values: SimpleNamespace(result=dialog.result))
        monkeypatch.setattr(app, "wait_window", lambda window: None)
        app.open_advanced_config_dialog()
        assert app.auto_confirm_score_var.get() == "80"
        state = app._collect_run_state()
        assert state["config"].auto_confirm_score == 80
        assert state["config"].tolerance_days == 5
        assert state["config"].dfs_date_window == 12
        assert build_matcher_config(default_gui_values()).tolerance_days == 31
        assert build_matcher_config(default_gui_values()).auto_confirm_score == 70
    finally:
        if dialog is not None and dialog.winfo_exists():
            dialog.destroy()
        app.destroy()
