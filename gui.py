"""
银行流水核对工具 v3.0 — 卡片化 GUI

布局结构（grid）：
  输入卡片 | 参数配置卡片 → 列映射卡片（可折叠）→ 执行卡片 → 日志卡片
"""

import threading
import queue
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any, Mapping, Sequence
from urllib.parse import urlsplit

import customtkinter as ctk
import pandas as pd

from data_structures import MatcherConfig
from data_loader import DataLoader, ParseErrorCollector
from input_precheck import InputPrecheckBlockedError, InputPrecheckReport
from validate import validate_config_params
from llm_assistant import LLMConfig, LLMAssistant
from application import run_reconciliation


# ─── 常量 ──────────────────────────────────────────────────────────────────────

# 按钮颜色
CLR_PRIMARY = "#2196F3"        # 主按钮 — 蓝色
CLR_PRIMARY_HOVER = "#1976D2"
CLR_SECONDARY = "#607D8B"     # 次按钮 — 蓝灰色
CLR_SECONDARY_HOVER = "#546E7A"
CLR_DANGER_BORDER = "#F44336" # 停止按钮 — 红色描边

# 卡片内边距
PAD_CARD = (14, 12)
PAD_CARD_TITLE = (0, 8)

# 组间间距
GAP_BETWEEN_CARDS = 10
MAIN_BUTTON = {"width": 150, "height": 40}
SECONDARY_BUTTON = {"width": 112, "height": 34}
CARD_RADIUS = 10
CONTENT_GAP = 12

AUXILIARY_COLUMN_KEYWORDS = (
    "摘要",
    "业务说明",
    "交易用途",
    "对方户名",
    "附言",
    "备注",
)
EXCLUDED_DEFAULT_AUXILIARY_KEYWORDS = (
    "账号",
    "帐号",
    "卡号",
    "凭证",
)


def default_gui_values() -> dict[str, str]:
    """返回界面初始值，正式业务名称不得缩写。"""
    return {
        "performance_materiality": "100000",
        "clearly_trivial_threshold": "5000",
        "auto_confirm_score": "70",
        "tolerance_days": "31",
        "dfs_window": "31",
        "dfs_depth": "30",
        "greedy_attempts": "3",
        "random_seed": "0",
        "similarity_threshold": "0.5",
        "similarity_high": "0.7",
        "max_candidates": "30",
        "memory_limit": "6.0",
        "batch_min_count": "10",
    }


def auto_select_auxiliary_columns(
    columns: Sequence[object],
) -> list[str]:
    """自动勾选业务文字列，账号、卡号和凭证号默认排除。"""
    selected = []
    for column in columns:
        label = str(column)
        if not any(keyword in label for keyword in AUXILIARY_COLUMN_KEYWORDS):
            continue
        if any(
            keyword in label
            for keyword in EXCLUDED_DEFAULT_AUXILIARY_KEYWORDS
        ):
            continue
        selected.append(label)
    return selected


def _auto_mapping_for_columns(
    columns: Sequence[object],
    *,
    is_bank: bool,
) -> dict[str, object]:
    """为新载入的文件生成可编辑的初始列映射。"""
    labels = [str(column) for column in columns]

    def find(keywords):
        for keyword in keywords:
            for label in labels:
                if keyword.lower() in label.strip().lower():
                    return label
        return None

    debit = find(
        ["借", "支出", "debit"]
        if is_bank
        else ["借方", "借", "debit"]
    )
    credit = find(
        ["贷", "收入", "credit"]
        if is_bank
        else ["贷方", "贷", "credit"]
    )
    amount = find(["金额", "amount", "发生额"])
    direction = find(["方向", "借贷标志"])
    if debit and credit and debit != credit:
        mode = "debit_credit"
    elif amount and direction:
        mode = "single_amount_with_direction"
    elif amount:
        mode = "signed_amount"
    else:
        mode = "debit_credit"

    return {
        "mode": mode,
        "date": find(["日期", "date", "交易时间", "time"]),
        "summary": find(
            ["摘要", "summary", "业务说明", "交易用途", "备注"]
        ),
        "voucher": find(["凭证", "voucher"]),
        "balance": find(["余额", "balance"]),
        "debit": debit,
        "credit": credit,
        "amount": amount,
        "direction": direction,
        "auxiliary_text_columns": auto_select_auxiliary_columns(labels),
    }


def _as_bool(value: object, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "是", "启用"}


def build_matcher_config(
    values: Mapping[str, object],
) -> MatcherConfig:
    """从界面快照构造匹配配置，主动忽略密钥等无关字段。"""
    defaults = default_gui_values()

    def text(key: str) -> str:
        value = values.get(key, defaults[key])
        return str(value).strip()

    try:
        return MatcherConfig(
            tolerance_days=int(text("tolerance_days")),
            dfs_date_window=int(text("dfs_window")),
            max_dfs_depth=int(text("dfs_depth")),
            allow_mixed_sign=_as_bool(
                values.get("allow_mixed_sign"),
                False,
            ),
            allow_zero_match=_as_bool(
                values.get("allow_zero_match"),
                False,
            ),
            allow_greedy_fallback=_as_bool(
                values.get("allow_greedy_fallback"),
                True,
            ),
            greedy_attempts=int(text("greedy_attempts")),
            random_seed=int(text("random_seed")),
            similarity_threshold=float(text("similarity_threshold")),
            similarity_high_threshold=float(text("similarity_high")),
            max_candidates=int(text("max_candidates")),
            memory_limit_gb=float(text("memory_limit")),
            performance_materiality=Decimal(
                text("performance_materiality")
            ),
            clearly_trivial_threshold=Decimal(
                text("clearly_trivial_threshold")
            ),
            auto_confirm_score=int(text("auto_confirm_score")),
            batch_min_count=int(text("batch_min_count")),
        )
    except (ValueError, InvalidOperation) as exc:
        raise ValueError("匹配参数中存在无法识别的数值") from exc


def build_llm_config(values: Mapping[str, object]) -> LLMConfig:
    """校验并构造只保存在当前会话内的大模型配置。"""
    enabled = _as_bool(values.get("enabled"), False)
    mode = str(values.get("mode", "online")).strip().lower() or "online"
    protocol = (
        str(values.get("protocol", "auto")).strip().lower() or "auto"
    )
    base_url = str(values.get("base_url", "")).strip().rstrip("/")
    model = str(values.get("model", "")).strip()
    api_key = str(values.get("api_key", ""))
    timeout_text = str(values.get("timeout_seconds", "")).strip()
    limit_text = str(values.get("candidate_limit", "")).strip()
    local_raw = values.get("local_fields", "")

    if not enabled:
        return LLMConfig(
            enabled=False,
            mode=mode,
            protocol=protocol,
            base_url=base_url,
            model=model,
            api_key=api_key,
        )
    if mode not in {"online", "local"}:
        raise ValueError("大模型模式必须为在线或本地 LM Studio")
    if protocol not in {"auto", "responses", "chat_completions"}:
        raise ValueError("大模型调用协议不受支持")
    if mode == "local" and not base_url:
        base_url = "http://localhost:1234/v1"
    parsed = urlsplit(base_url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("大模型服务地址格式不正确")
    if not model:
        raise ValueError("启用大模型时必须选择模型")
    try:
        timeout = float(timeout_text or "30")
    except ValueError as exc:
        raise ValueError("大模型超时必须是数字") from exc
    if not 1 <= timeout <= 300:
        raise ValueError("大模型超时必须在1到300秒之间")
    try:
        candidate_limit = int(limit_text or "5")
    except ValueError as exc:
        raise ValueError("大模型候选数必须是整数") from exc
    if not 1 <= candidate_limit <= 20:
        raise ValueError("大模型候选数必须在1到20之间")

    if isinstance(local_raw, str):
        normalized = local_raw.replace("，", ",").replace("；", ",")
        local_fields = tuple(
            field.strip()
            for field in normalized.split(",")
            if field.strip()
        )
    elif isinstance(local_raw, Sequence):
        local_fields = tuple(
            str(field).strip()
            for field in local_raw
            if str(field).strip()
        )
    else:
        local_fields = ()
    return LLMConfig(
        enabled=True,
        mode=mode,
        protocol=protocol,
        base_url=base_url,
        model=model,
        api_key=api_key,
        timeout_seconds=timeout,
        candidate_limit=candidate_limit,
        local_fields=local_fields,
    )


# ─── ColumnMappingFrame ────────────────────────────────────────────────────────

class ColumnMappingFrame(ctk.CTkFrame):
    """列映射卡片内部的映射组件，包含模式选择与列下拉框"""

    def __init__(self, parent, title, is_bank=False):
        super().__init__(parent, fg_color="transparent")
        self.vars = {}
        self.aux_vars = {}
        self.columns = []
        self.is_bank = is_bank

        # 标题
        ctk.CTkLabel(self, text=title, font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 4)
        )

        # 模式选择
        mode_row = ctk.CTkFrame(self, fg_color="transparent")
        mode_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ctk.CTkLabel(mode_row, text="模式:", width=60).pack(side=ctk.LEFT)
        self.mode_var = ctk.StringVar(value="借贷分列")
        ctk.CTkComboBox(
            mode_row, variable=self.mode_var,
            values=["借贷分列", "单列金额+方向", "单列金额(含正负)"],
            width=180, command=self.refresh, state="readonly"
        ).pack(side=ctk.LEFT, padx=(4, 0))

        # 映射字段容器
        self.grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_frame.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.grid_frame.grid_columnconfigure(0, weight=0)
        self.grid_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self,
            text="匹配辅助文字（可多选）",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(
            row=3,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(8, 2),
        )
        self.aux_frame = ctk.CTkScrollableFrame(
            self,
            height=92,
            fg_color=("gray95", "gray17"),
        )
        self.aux_frame.grid(
            row=4,
            column=0,
            columnspan=2,
            sticky="ew",
        )
        self.aux_frame.grid_columnconfigure(0, weight=1)
        self.aux_frame.grid_columnconfigure(1, weight=1)
        self.create_widgets()
        self.create_auxiliary_widgets()

    def set_columns(self, cols):
        self.columns = cols
        self.refresh()

    def refresh(self, _=None):
        self.create_widgets()
        self.create_auxiliary_widgets()

    def create_auxiliary_widgets(self):
        for widget in self.aux_frame.winfo_children():
            widget.destroy()
        self.aux_vars = {}
        auto_selected = set(
            auto_select_auxiliary_columns(self.columns)
        )
        if not self.columns:
            ctk.CTkLabel(
                self.aux_frame,
                text="选择文件后显示可选文字列",
                text_color=("gray45", "gray65"),
            ).grid(row=0, column=0, sticky="w", padx=4, pady=2)
            return
        for index, column in enumerate(self.columns):
            label = str(column)
            variable = ctk.BooleanVar(value=label in auto_selected)
            self.aux_vars[label] = variable
            ctk.CTkCheckBox(
                self.aux_frame,
                text=label,
                variable=variable,
                width=150,
            ).grid(
                row=index // 2,
                column=index % 2,
                sticky="w",
                padx=4,
                pady=2,
            )

    def create_widgets(self):
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.vars = {}

        row_idx = 0

        def add(label, key, allow_none=False):
            nonlocal row_idx
            ctk.CTkLabel(self.grid_frame, text=label, width=80, anchor="w").grid(
                row=row_idx, column=0, sticky="w", pady=2
            )
            vals = ["(无)"] + self.columns if allow_none else self.columns
            var = ctk.StringVar(value=vals[0] if vals else "")
            ctk.CTkComboBox(self.grid_frame, variable=var, values=vals, state="readonly").grid(
                row=row_idx, column=1, sticky="ew", padx=(4, 0), pady=2
            )
            self.vars[key] = var

            # 自动匹配列名
            normalized_cols = {str(c).strip().lower(): c for c in self.columns}
            keywords = {
                "date": ["日期", "date", "交易时间", "time"],
                "summary": ["摘要", "summary", "备注", "对方户名", "用途", "业务说明"],
                "voucher": ["凭证", "voucher"],
                "balance": ["余额", "balance"],
                "debit": ["借", "debit", "支"],
                "credit": ["贷", "credit", "收"],
                "amount": ["金额", "amount", "发生额"],
                "direction": ["方向", "借贷标志"],
            }
            candidates = keywords.get(key, [])
            if "借" in label:
                candidates = keywords["debit"]
            if "贷" in label:
                candidates = keywords["credit"]

            for kw in candidates:
                for col_norm, col_orig in normalized_cols.items():
                    if kw in col_norm:
                        var.set(col_orig)
                        row_idx += 1
                        return

            row_idx += 1

        add("日期", "date")
        add("摘要", "summary")
        add("凭证", "voucher", True)
        add("余额", "balance", True)

        mode = self.mode_var.get()
        if mode == "借贷分列":
            add("借/支" if self.is_bank else "借方", "debit")
            add("贷/收" if self.is_bank else "贷方", "credit")
        elif mode == "单列金额+方向":
            add("金额", "amount")
            add("方向", "direction")
        else:
            add("金额", "amount")

    def get_mapping(self):
        m = {k: v.get() for k, v in self.vars.items()}
        for k, v in m.items():
            if v == "(无)":
                m[k] = None
        mode = self.mode_var.get()
        m["mode"] = (
            "debit_credit"
            if mode == "借贷分列"
            else ("single_amount_with_direction" if mode == "单列金额+方向" else "signed_amount")
        )
        m["auxiliary_text_columns"] = [
            column
            for column, variable in self.aux_vars.items()
            if variable.get()
        ]
        return m

    def set_mapping(self, mapping):
        """恢复已保存的列映射，取消窗口时主界面状态不会变化。"""
        if not mapping:
            return
        mode_names = {
            "debit_credit": "借贷分列",
            "single_amount_with_direction": "单列金额+方向",
            "signed_amount": "单列金额(含正负)",
        }
        self.mode_var.set(
            mode_names.get(mapping.get("mode"), "借贷分列")
        )
        self.refresh()
        for key, value in mapping.items():
            if key in self.vars and value in self.columns:
                self.vars[key].set(value)
        if "auxiliary_text_columns" in mapping:
            selected_auxiliary = set(
                mapping.get("auxiliary_text_columns", [])
            )
            for column, variable in self.aux_vars.items():
                variable.set(column in selected_auxiliary)


class ColumnMappingDialog(ctk.CTkToplevel):
    """独立可缩放的核心列映射与辅助字段多选窗口。"""

    def __init__(
        self,
        parent,
        bank_columns,
        journal_columns,
        bank_mapping,
        journal_mapping,
    ):
        super().__init__(parent)
        self.title("配置列映射与辅助字段")
        self.geometry("1080x720")
        self.minsize(900, 620)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(
            self,
            corner_radius=10,
            border_width=1,
            border_color=("#D8DEE9", "#3A3F4B"),
        )
        header.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=14,
            pady=(14, 8),
        )
        ctk.CTkLabel(
            header,
            text="配置列映射与辅助字段",
            font=ctk.CTkFont(size=17, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(
            header,
            text=(
                "账号等敏感字段可以用于本机比较；"
                "在线大模型不会发送其原值。"
            ),
            text_color=("gray40", "gray70"),
        ).grid(row=1, column=0, sticky="w", padx=14, pady=(0, 12))

        content = ctk.CTkScrollableFrame(self, corner_radius=10)
        content.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=14,
            pady=8,
        )
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=1)
        self.bank_frame = ColumnMappingFrame(
            content,
            "银行流水",
            is_bank=True,
        )
        self.bank_frame.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(6, 10),
            pady=6,
        )
        self.journal_frame = ColumnMappingFrame(
            content,
            "日记账",
            is_bank=False,
        )
        self.journal_frame.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(10, 6),
            pady=6,
        )
        self.bank_frame.set_columns(list(bank_columns))
        self.journal_frame.set_columns(list(journal_columns))
        self.bank_frame.set_mapping(bank_mapping)
        self.journal_frame.set_mapping(journal_mapping)

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(
            row=2,
            column=0,
            sticky="e",
            padx=14,
            pady=(8, 14),
        )
        ctk.CTkButton(
            footer,
            text="取消",
            width=100,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray85"),
            command=self._cancel,
        ).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(
            footer,
            text="保存列映射",
            width=140,
            fg_color=CLR_PRIMARY,
            hover_color=CLR_PRIMARY_HOVER,
            command=self._save,
        ).grid(row=0, column=1)

    def _save(self):
        self.result = (
            self.bank_frame.get_mapping(),
            self.journal_frame.get_mapping(),
        )
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class LLMConfigDialog(ctk.CTkToplevel):
    """大模型会话配置窗口，密钥不写入任何配置文件。"""

    def __init__(self, parent, current: LLMConfig):
        super().__init__(parent)
        self.title("大模型辅助配置")
        self.geometry("650x610")
        self.minsize(620, 560)
        self.transient(parent)
        self.grab_set()
        self.result = None

        self.enabled_var = ctk.BooleanVar(value=current.enabled)
        self.mode_var = ctk.StringVar(value=current.mode or "online")
        self.protocol_var = ctk.StringVar(
            value=current.protocol or "auto"
        )
        self.base_url_var = ctk.StringVar(value=current.base_url)
        self.model_var = ctk.StringVar(value=current.model)
        self.api_key_var = ctk.StringVar(value=current.api_key)
        self.timeout_var = ctk.StringVar(
            value=str(current.timeout_seconds or 30)
        )
        self.limit_var = ctk.StringVar(
            value=str(current.candidate_limit or 5)
        )
        self.local_fields_var = ctk.StringVar(
            value=", ".join(current.local_fields)
        )
        self.status_var = ctk.StringVar(
            value="大模型是可选增强项；关闭时不影响基础核对。"
        )
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        header = ctk.CTkFrame(
            self,
            corner_radius=10,
            border_width=1,
            border_color=("#D8DEE9", "#3A3F4B"),
        )
        header.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=14,
            pady=(14, 8),
        )
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header,
            text="大模型语义辅助",
            font=ctk.CTkFont(size=17, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(
            header,
            text="只比较已有候选的文字语义，不能覆盖金额和重要性规则。",
            text_color=("gray40", "gray70"),
        ).grid(row=1, column=0, sticky="w", padx=14, pady=(0, 12))

        body = ctk.CTkScrollableFrame(self, corner_radius=10)
        body.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=14,
            pady=8,
        )
        body.grid_columnconfigure(1, weight=1)
        ctk.CTkCheckBox(
            body,
            text="启用大模型辅助",
            variable=self.enabled_var,
        ).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="w",
            padx=12,
            pady=(12, 8),
        )

        fields = [
            ("运行方式", "mode"),
            ("调用协议", "protocol"),
            ("服务地址", "base_url"),
            ("模型", "model"),
            ("API 密钥", "api_key"),
            ("超时（秒）", "timeout"),
            ("每组候选上限", "limit"),
            ("本地可发送字段", "local_fields"),
        ]
        for row_index, (label, key) in enumerate(fields, start=1):
            ctk.CTkLabel(
                body,
                text=label,
                width=120,
                anchor="w",
            ).grid(
                row=row_index,
                column=0,
                sticky="w",
                padx=(12, 8),
                pady=5,
            )
            if key == "mode":
                widget = ctk.CTkComboBox(
                    body,
                    variable=self.mode_var,
                    values=["online", "local"],
                    state="readonly",
                    command=self._mode_changed,
                )
            elif key == "protocol":
                widget = ctk.CTkComboBox(
                    body,
                    variable=self.protocol_var,
                    values=["auto", "responses", "chat_completions"],
                    state="readonly",
                )
            elif key == "model":
                self.model_combo = ctk.CTkComboBox(
                    body,
                    variable=self.model_var,
                    values=[self.model_var.get()]
                    if self.model_var.get()
                    else [""],
                )
                widget = self.model_combo
            elif key == "api_key":
                widget = ctk.CTkEntry(
                    body,
                    textvariable=self.api_key_var,
                    show="●",
                    placeholder_text="仅保存在当前程序内存",
                )
            else:
                variable = {
                    "base_url": self.base_url_var,
                    "timeout": self.timeout_var,
                    "limit": self.limit_var,
                    "local_fields": self.local_fields_var,
                }[key]
                widget = ctk.CTkEntry(
                    body,
                    textvariable=variable,
                )
            widget.grid(
                row=row_index,
                column=1,
                sticky="ew",
                padx=(0, 12),
                pady=5,
            )

        ctk.CTkLabel(
            body,
            text=(
                "在线：日期、金额、摘要、业务说明、交易用途、附言、备注＋"
                "本机生成的敏感字段相同/不同结论；不发送账号、卡号等原值。"
                "本地：只发送上方填写字段。"
            ),
            justify="left",
            wraplength=560,
            text_color=("gray40", "gray70"),
        ).grid(
            row=len(fields) + 1,
            column=0,
            columnspan=2,
            sticky="w",
            padx=12,
            pady=(8, 4),
        )
        action_row = ctk.CTkFrame(body, fg_color="transparent")
        action_row.grid(
            row=len(fields) + 2,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=12,
            pady=8,
        )
        ctk.CTkButton(
            action_row,
            text="刷新模型",
            width=110,
            fg_color=CLR_SECONDARY,
            hover_color=CLR_SECONDARY_HOVER,
            command=self._refresh_models,
        ).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(
            action_row,
            text="测试连接",
            width=110,
            fg_color=CLR_SECONDARY,
            hover_color=CLR_SECONDARY_HOVER,
            command=self._test_connection,
        ).grid(row=0, column=1)
        ctk.CTkLabel(
            body,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            wraplength=520,
        ).grid(
            row=len(fields) + 3,
            column=0,
            columnspan=2,
            sticky="ew",
            padx=12,
            pady=(2, 12),
        )

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(
            row=2,
            column=0,
            sticky="e",
            padx=14,
            pady=(8, 14),
        )
        ctk.CTkButton(
            footer,
            text="取消",
            width=100,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray85"),
            command=self._cancel,
        ).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(
            footer,
            text="保存当前会话配置",
            width=170,
            fg_color=CLR_PRIMARY,
            hover_color=CLR_PRIMARY_HOVER,
            command=self._save,
        ).grid(row=0, column=1)

    def _mode_changed(self, _value=None):
        if self.mode_var.get() == "local" and not self.base_url_var.get():
            self.base_url_var.set("http://localhost:1234/v1")

    def _values(self, *, force_enabled=None) -> dict[str, object]:
        return {
            "enabled": (
                self.enabled_var.get()
                if force_enabled is None
                else force_enabled
            ),
            "mode": self.mode_var.get(),
            "protocol": self.protocol_var.get(),
            "base_url": self.base_url_var.get(),
            "model": self.model_var.get(),
            "api_key": self.api_key_var.get(),
            "timeout_seconds": self.timeout_var.get(),
            "candidate_limit": self.limit_var.get(),
            "local_fields": self.local_fields_var.get(),
        }

    def _connection_config(self) -> LLMConfig:
        values = self._values(force_enabled=True)
        if not str(values["model"]).strip():
            values["model"] = "__connection_test__"
        return build_llm_config(values)

    def _refresh_models(self):
        self.status_var.set("正在读取模型列表…")

        def worker():
            try:
                assistant = LLMAssistant(self._connection_config())
                models = assistant.list_models()
                if not models:
                    raise ValueError(
                        assistant.last_error or "服务没有返回模型"
                    )
                self.after(0, lambda: self._show_models(models))
            except Exception as exc:
                self.after(
                    0,
                    lambda: self.status_var.set(f"读取失败：{exc}"),
                )

        threading.Thread(target=worker, daemon=True).start()

    def _show_models(self, models):
        self.model_combo.configure(values=models)
        if self.model_var.get() not in models:
            self.model_var.set(models[0])
        self.status_var.set(f"已读取 {len(models)} 个模型。")

    def _test_connection(self):
        self.status_var.set("正在测试连接…")

        def worker():
            try:
                assistant = LLMAssistant(self._connection_config())
                success, message = assistant.test_connection()
            except Exception as exc:
                success, message = False, str(exc)
            self.after(
                0,
                lambda: self.status_var.set(
                    ("连接成功：" if success else "连接失败：")
                    + message
                ),
            )

        threading.Thread(target=worker, daemon=True).start()

    def _save(self):
        try:
            self.result = build_llm_config(self._values())
        except ValueError as exc:
            messagebox.showerror("大模型配置错误", str(exc), parent=self)
            return
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class AdvancedConfigDialog(ctk.CTkToplevel):
    """不挤占主窗口空间的高级匹配参数窗口。"""

    def __init__(self, parent, current_values):
        super().__init__(parent)
        self.title("高级匹配参数")
        self.geometry("620x650")
        self.minsize(560, 560)
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.values = dict(current_values)
        self.variables = {
            key: ctk.StringVar(value=str(value))
            for key, value in self.values.items()
            if key not in {
                "allow_mixed_sign",
                "allow_zero_match",
                "allow_greedy_fallback",
            }
        }
        self.allow_mixed_sign_var = ctk.BooleanVar(
            value=_as_bool(self.values.get("allow_mixed_sign"))
        )
        self.allow_zero_match_var = ctk.BooleanVar(
            value=_as_bool(self.values.get("allow_zero_match"))
        )
        self.allow_greedy_var = ctk.BooleanVar(
            value=_as_bool(
                self.values.get("allow_greedy_fallback"),
                True,
            )
        )
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(
            self,
            text="高级匹配参数",
            font=ctk.CTkFont(size=17, weight="bold"),
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=18,
            pady=(16, 8),
        )
        body = ctk.CTkScrollableFrame(self, corner_radius=10)
        body.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=14,
            pady=8,
        )
        body.grid_columnconfigure(1, weight=1)
        field_labels = [
            ("日期容差（天）", "tolerance_days"),
            ("组合窗口（天）", "dfs_window"),
            ("组合最大深度", "dfs_depth"),
            ("批量最少笔数", "batch_min_count"),
            ("贪心尝试次数", "greedy_attempts"),
            ("随机种子（-1为随机）", "random_seed"),
            ("文字相似度阈值", "similarity_threshold"),
            ("高相似度阈值", "similarity_high"),
            ("每笔最大候选数", "max_candidates"),
            ("内存限制（GB）", "memory_limit"),
        ]
        for row_index, (label, key) in enumerate(field_labels):
            ctk.CTkLabel(
                body,
                text=label,
                width=180,
                anchor="w",
            ).grid(
                row=row_index,
                column=0,
                sticky="w",
                padx=(12, 8),
                pady=5,
            )
            ctk.CTkEntry(
                body,
                textvariable=self.variables[key],
            ).grid(
                row=row_index,
                column=1,
                sticky="ew",
                padx=(0, 12),
                pady=5,
            )
        checkbox_row = len(field_labels)
        ctk.CTkCheckBox(
            body,
            text="允许异号组合",
            variable=self.allow_mixed_sign_var,
        ).grid(
            row=checkbox_row,
            column=0,
            sticky="w",
            padx=12,
            pady=(12, 5),
        )
        ctk.CTkCheckBox(
            body,
            text="允许零金额匹配",
            variable=self.allow_zero_match_var,
        ).grid(
            row=checkbox_row,
            column=1,
            sticky="w",
            padx=12,
            pady=(12, 5),
        )
        ctk.CTkCheckBox(
            body,
            text="启用贪心兜底",
            variable=self.allow_greedy_var,
        ).grid(
            row=checkbox_row + 1,
            column=0,
            sticky="w",
            padx=12,
            pady=5,
        )
        ctk.CTkLabel(
            body,
            text="日期显示格式",
            anchor="w",
        ).grid(
            row=checkbox_row + 2,
            column=0,
            sticky="w",
            padx=12,
            pady=5,
        )
        date_formats = [
            "auto", "YYYY-MM-DD", "YYYY/MM/DD", "YYYYMMDD",
            "DD-MM-YYYY", "DD/MM/YYYY", "MM/DD/YYYY",
            "DD.MM.YYYY", "YYYY.MM.DD",
        ]
        ctk.CTkComboBox(
            body,
            variable=self.variables["date_format"],
            values=date_formats,
            state="readonly",
        ).grid(
            row=checkbox_row + 2,
            column=1,
            sticky="ew",
            padx=(0, 12),
            pady=5,
        )

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(
            row=2,
            column=0,
            sticky="e",
            padx=14,
            pady=(8, 14),
        )
        ctk.CTkButton(
            footer,
            text="取消",
            width=100,
            fg_color="transparent",
            border_width=1,
            text_color=("gray20", "gray85"),
            command=self._cancel,
        ).grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(
            footer,
            text="保存高级参数",
            width=150,
            fg_color=CLR_PRIMARY,
            hover_color=CLR_PRIMARY_HOVER,
            command=self._save,
        ).grid(row=0, column=1)

    def _save(self):
        updated = dict(self.values)
        for key, variable in self.variables.items():
            updated[key] = variable.get()
        updated["allow_mixed_sign"] = self.allow_mixed_sign_var.get()
        updated["allow_zero_match"] = self.allow_zero_match_var.get()
        updated["allow_greedy_fallback"] = self.allow_greedy_var.get()
        try:
            build_matcher_config(updated)
        except ValueError as exc:
            messagebox.showerror("高级参数错误", str(exc), parent=self)
            return
        self.result = updated
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ─── 辅助：创建卡片容器 ─────────────────────────────────────────────────────────

def _make_card(parent, title, row, column, columnspan=1, rowspan=1, sticky="nsew", **kw):
    """创建一个带标题的卡片 CTkFrame，返回 (card_frame, title_label)"""
    card = ctk.CTkFrame(parent, corner_radius=CARD_RADIUS, border_width=1, border_color=("#E0E0E0", "#3A3A3A"))
    card.grid(
        row=row, column=column, columnspan=columnspan, rowspan=rowspan,
        sticky=sticky, padx=4, pady=GAP_BETWEEN_CARDS, **kw
    )
    # 卡片标题
    title_label = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold"))
    title_label.grid(row=0, column=0, sticky="w", padx=PAD_CARD[0], pady=(PAD_CARD[1], PAD_CARD_TITLE[1]))
    return card, title_label


# ─── ReconciliationApp ─────────────────────────────────────────────────────────

class ReconciliationApp(ctk.CTk):
    """银行流水核对工具主窗口 — 卡片化 grid 布局"""

    def __init__(self):
        super().__init__()
        self.title("银行流水核对工具 v3.0")
        self.geometry("1200x850")
        self.minsize(960, 700)

        self.error_collector = ParseErrorCollector()
        self.loader = DataLoader(logger=self.log, error_collector=self.error_collector)
        self.log_queue = queue.Queue()
        self.log_lock = threading.Lock()
        self.after(100, self.process_log_queue)
        self.matcher = None
        self.theme_mode = "system"
        self.llm_config = LLMConfig()
        self.bank_columns = []
        self.journal_columns = []
        self.bank_mapping_state = {}
        self.journal_mapping_state = {}

        self._init_ui()

    # ─── UI 构建 ──────────────────────────────────────────────────────────

    def _init_ui(self):
        # 根窗口 grid 配置：紧凑设置、执行、日志
        self.grid_rowconfigure(0, weight=0)  # 输入 + 策略 + 配置入口
        self.grid_rowconfigure(1, weight=0)  # 执行
        self.grid_rowconfigure(2, weight=1)  # 日志 — 占满剩余空间
        self.grid_columnconfigure(0, weight=1)

        self._build_input_and_config_cards()
        self._build_action_card()
        self._build_log_card()

    # ─── 输入卡片 + 参数配置卡片（并排） ─────────────────────────────────

    def _build_input_and_config_cards(self):
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.grid(row=0, column=0, sticky="ew", padx=4, pady=GAP_BETWEEN_CARDS)
        container.grid_columnconfigure(0, weight=3)
        container.grid_columnconfigure(1, weight=2)
        container.grid_columnconfigure(2, weight=2)

        # ── 左侧：输入卡片 ──
        self._build_input_card(container)

        # ── 中间：核心审计策略 ──
        self._build_config_card(container)

        # ── 右侧：配置入口 ──
        self._build_matching_and_model_card(container)

    def _build_input_card(self, parent):
        card, _ = _make_card(parent, "输入文件", row=0, column=0)

        self.bank_path = ctk.StringVar()
        self.journal_path = ctk.StringVar()
        self.bank_skip = ctk.StringVar(value="0")
        self.journal_skip = ctk.StringVar(value="0")
        self.bank_header_rows = ctk.StringVar(value="1")
        self.journal_header_rows = ctk.StringVar(value="1")

        # 银行流行
        self._build_file_row(
            card, 1, "银行流水:", self.bank_path, self.bank_skip,
            self.bank_header_rows,
                             lambda: self.browse("bank"), lambda: self.auto_detect("bank"))

        # 日记账行
        self._build_file_row(
            card, 2, "日记账:", self.journal_path, self.journal_skip,
            self.journal_header_rows,
                             lambda: self.browse("journal"), lambda: self.auto_detect("journal"))

        # 配置列自适应
        card.grid_columnconfigure(1, weight=1)

    def _build_file_row(
        self,
        card,
        row,
        label,
        path_var,
        skip_var,
        header_rows_var,
        browse_cmd,
        detect_cmd,
    ):
        ctk.CTkLabel(card, text=label, width=70, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(PAD_CARD[0], 4), pady=3
        )
        ctk.CTkEntry(card, textvariable=path_var).grid(
            row=row, column=1, sticky="ew", padx=2, pady=3
        )
        ctk.CTkButton(
            card, text="浏览", width=60,
            fg_color=CLR_SECONDARY, hover_color=CLR_SECONDARY_HOVER,
            command=browse_cmd
        ).grid(row=row, column=2, padx=2, pady=3)

        ctk.CTkLabel(card, text="跳过:", width=40, anchor="e").grid(
            row=row, column=3, sticky="e", padx=(8, 0), pady=3
        )
        ctk.CTkEntry(card, textvariable=skip_var, width=45).grid(
            row=row, column=4, padx=2, pady=3
        )
        ctk.CTkLabel(card, text="表头行数:", width=64, anchor="e").grid(
            row=row, column=5, sticky="e", padx=(6, 0), pady=3
        )
        ctk.CTkEntry(card, textvariable=header_rows_var, width=38).grid(
            row=row, column=6, padx=2, pady=3
        )
        ctk.CTkButton(
            card, text="自动检测", width=72,
            fg_color=CLR_SECONDARY, hover_color=CLR_SECONDARY_HOVER,
            command=detect_cmd
        ).grid(row=row, column=7, padx=(2, PAD_CARD[0]), pady=3)

    def _build_config_card(self, parent):
        card, _ = _make_card(parent, "审计策略", row=0, column=1)

        # 参数变量
        defaults = default_gui_values()
        self.performance_materiality_var = ctk.StringVar(
            value=defaults["performance_materiality"]
        )
        self.clearly_trivial_threshold_var = ctk.StringVar(
            value=defaults["clearly_trivial_threshold"]
        )
        self.auto_confirm_score_var = ctk.StringVar(
            value=defaults["auto_confirm_score"]
        )
        self.tolerance_days_var = ctk.StringVar(
            value=defaults["tolerance_days"]
        )
        self.dfs_window_var = ctk.StringVar(value=defaults["dfs_window"])
        self.dfs_depth_var = ctk.StringVar(value=defaults["dfs_depth"])
        self.allow_mixed_sign_var = ctk.BooleanVar(value=False)
        self.allow_zero_match_var = ctk.BooleanVar(value=False)
        self.allow_greedy_var = ctk.BooleanVar(value=True)
        self.greedy_attempts_var = ctk.StringVar(
            value=defaults["greedy_attempts"]
        )
        self.random_seed_var = ctk.StringVar(
            value=defaults["random_seed"]
        )
        self.similarity_threshold_var = ctk.StringVar(
            value=defaults["similarity_threshold"]
        )
        self.similarity_high_var = ctk.StringVar(
            value=defaults["similarity_high"]
        )
        self.max_candidates_var = ctk.StringVar(
            value=defaults["max_candidates"]
        )
        self.memory_limit_var = ctk.StringVar(
            value=defaults["memory_limit"]
        )
        self.batch_min_count_var = ctk.StringVar(
            value=defaults["batch_min_count"]
        )
        self.date_fmt_var = ctk.StringVar(value="auto")
        self.llm_status_var = ctk.StringVar(value="大模型辅助：关闭")

        # 首页只保留用户确认过的三个核心策略
        params = [
            ("实际执行重要性水平", self.performance_materiality_var),
            ("明显微小错报临界值", self.clearly_trivial_threshold_var),
            ("自动确认最低综合可信度", self.auto_confirm_score_var),
        ]

        for i, (label, var) in enumerate(params):
            r = i + 1
            ctk.CTkLabel(
                card,
                text=label,
                anchor="w",
            ).grid(
                row=r,
                column=0,
                sticky="w",
                padx=(PAD_CARD[0], 6),
                pady=5,
            )
            ctk.CTkEntry(
                card,
                textvariable=var,
                width=100,
            ).grid(
                row=r,
                column=1,
                sticky="e",
                padx=(0, PAD_CARD[0]),
                pady=5,
            )
        ctk.CTkLabel(
            card,
            text="超过重要性水平的组始终进入人工复核。",
            text_color=("gray40", "gray70"),
            wraplength=260,
            justify="left",
        ).grid(
            row=len(params) + 1,
            column=0,
            columnspan=2,
            sticky="w",
            padx=PAD_CARD[0],
            pady=(8, PAD_CARD[1]),
        )
        card.grid_columnconfigure(0, weight=1)

    def _build_matching_and_model_card(self, parent):
        card, _ = _make_card(
            parent,
            "匹配设置与增强",
            row=0,
            column=2,
        )
        card.grid_columnconfigure(0, weight=1)
        self.mapping_status_var = ctk.StringVar(
            value="请先选择文件并读取列"
        )

        ctk.CTkLabel(
            card,
            text="列映射与辅助字段",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=PAD_CARD[0],
            pady=(2, 3),
        )
        ctk.CTkButton(
            card,
            text="配置列映射与辅助字段",
            fg_color=CLR_SECONDARY,
            hover_color=CLR_SECONDARY_HOVER,
            command=self.open_column_mapping_dialog,
            **SECONDARY_BUTTON,
        ).grid(
            row=2,
            column=0,
            sticky="ew",
            padx=PAD_CARD[0],
            pady=3,
        )
        ctk.CTkLabel(
            card,
            textvariable=self.mapping_status_var,
            text_color=("gray40", "gray70"),
            anchor="w",
        ).grid(
            row=3,
            column=0,
            sticky="ew",
            padx=PAD_CARD[0],
            pady=(1, 8),
        )
        ctk.CTkButton(
            card,
            text="高级匹配参数",
            fg_color=CLR_SECONDARY,
            hover_color=CLR_SECONDARY_HOVER,
            command=self.open_advanced_config_dialog,
            **SECONDARY_BUTTON,
        ).grid(
            row=4,
            column=0,
            sticky="ew",
            padx=PAD_CARD[0],
            pady=3,
        )

        ctk.CTkLabel(
            card,
            text="大模型辅助（可选）",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).grid(
            row=5,
            column=0,
            sticky="w",
            padx=PAD_CARD[0],
            pady=(12, 3),
        )
        ctk.CTkButton(
            card,
            text="配置与测试",
            fg_color=CLR_SECONDARY,
            hover_color=CLR_SECONDARY_HOVER,
            command=self.open_llm_config_dialog,
            **SECONDARY_BUTTON,
        ).grid(
            row=6,
            column=0,
            sticky="ew",
            padx=PAD_CARD[0],
            pady=3,
        )
        ctk.CTkLabel(
            card,
            textvariable=self.llm_status_var,
            text_color=("gray40", "gray70"),
            anchor="w",
            wraplength=250,
            justify="left",
        ).grid(
            row=7,
            column=0,
            sticky="ew",
            padx=PAD_CARD[0],
            pady=(1, PAD_CARD[1]),
        )

    def open_column_mapping_dialog(self):
        if not self.bank_columns or not self.journal_columns:
            messagebox.showinfo(
                "列映射",
                "请先选择银行流水和日记账文件，并完成列读取。",
            )
            return
        dialog = ColumnMappingDialog(
            self,
            self.bank_columns,
            self.journal_columns,
            self.bank_mapping_state,
            self.journal_mapping_state,
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return
        (
            self.bank_mapping_state,
            self.journal_mapping_state,
        ) = dialog.result
        bank_aux = len(
            self.bank_mapping_state.get("auxiliary_text_columns", [])
        )
        journal_aux = len(
            self.journal_mapping_state.get(
                "auxiliary_text_columns",
                [],
            )
        )
        self.mapping_status_var.set(
            f"已配置：银行{bank_aux}列，日记账{journal_aux}列"
        )

    def _current_advanced_values(self):
        return {
            **default_gui_values(),
            "performance_materiality": (
                self.performance_materiality_var.get()
            ),
            "clearly_trivial_threshold": (
                self.clearly_trivial_threshold_var.get()
            ),
            "auto_confirm_score": self.auto_confirm_score_var.get(),
            "tolerance_days": self.tolerance_days_var.get(),
            "dfs_window": self.dfs_window_var.get(),
            "dfs_depth": self.dfs_depth_var.get(),
            "batch_min_count": self.batch_min_count_var.get(),
            "greedy_attempts": self.greedy_attempts_var.get(),
            "random_seed": self.random_seed_var.get(),
            "similarity_threshold": self.similarity_threshold_var.get(),
            "similarity_high": self.similarity_high_var.get(),
            "max_candidates": self.max_candidates_var.get(),
            "memory_limit": self.memory_limit_var.get(),
            "date_format": self.date_fmt_var.get(),
            "allow_mixed_sign": self.allow_mixed_sign_var.get(),
            "allow_zero_match": self.allow_zero_match_var.get(),
            "allow_greedy_fallback": self.allow_greedy_var.get(),
        }

    def open_advanced_config_dialog(self):
        dialog = AdvancedConfigDialog(
            self,
            self._current_advanced_values(),
        )
        self.wait_window(dialog)
        if dialog.result is None:
            return
        values = dialog.result
        variable_map = {
            "tolerance_days": self.tolerance_days_var,
            "dfs_window": self.dfs_window_var,
            "dfs_depth": self.dfs_depth_var,
            "batch_min_count": self.batch_min_count_var,
            "greedy_attempts": self.greedy_attempts_var,
            "random_seed": self.random_seed_var,
            "similarity_threshold": self.similarity_threshold_var,
            "similarity_high": self.similarity_high_var,
            "max_candidates": self.max_candidates_var,
            "memory_limit": self.memory_limit_var,
            "date_format": self.date_fmt_var,
        }
        for key, variable in variable_map.items():
            variable.set(str(values[key]))
        self.allow_mixed_sign_var.set(
            _as_bool(values.get("allow_mixed_sign"))
        )
        self.allow_zero_match_var.set(
            _as_bool(values.get("allow_zero_match"))
        )
        self.allow_greedy_var.set(
            _as_bool(
                values.get("allow_greedy_fallback"),
                True,
            )
        )

    def open_llm_config_dialog(self):
        dialog = LLMConfigDialog(self, self.llm_config)
        self.wait_window(dialog)
        if dialog.result is None:
            return
        self.llm_config = dialog.result
        if self.llm_config.enabled:
            mode_name = (
                "本地 LM Studio"
                if self.llm_config.mode == "local"
                else "在线 API"
            )
            self.llm_status_var.set(
                f"大模型辅助：已启用（{mode_name}）"
            )
        else:
            self.llm_status_var.set("大模型辅助：关闭")

    # ─── 执行卡片 ─────────────────────────────────────────────────────────

    def _build_action_card(self):
        card, _ = _make_card(self, "", row=1, column=0)

        # 主按钮：开始核对
        self.btn_start = ctk.CTkButton(
            card, text="▶  开始核对",
            fg_color=CLR_PRIMARY, hover_color=CLR_PRIMARY_HOVER,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.start,
            **MAIN_BUTTON,
        )
        self.btn_start.grid(row=0, column=0, sticky="w", padx=PAD_CARD[0], pady=PAD_CARD[1])

        progress_frame = ctk.CTkFrame(card, fg_color="transparent")
        progress_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=CONTENT_GAP,
            pady=8,
        )
        progress_frame.grid_columnconfigure(0, weight=1)
        self.status_var = ctk.StringVar(value="准备就绪")
        ctk.CTkLabel(
            progress_frame,
            textvariable=self.status_var,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 3))
        self.prog = ctk.CTkProgressBar(progress_frame)
        self.prog.grid(row=1, column=0, sticky="ew")
        self.prog.set(0)

        # 弱按钮：停止（红色描边）
        self.btn_stop = ctk.CTkButton(
            card, text="■  停止", width=100, height=40,
            fg_color="transparent", border_color=CLR_DANGER_BORDER,
            border_width=2, text_color=("#F44336", "#F44336"),
            hover_color=("#FFCDD2", "#5C3030"),
            state="disabled", command=self.stop_process
        )
        self.btn_stop.grid(row=0, column=2, sticky="e", padx=(CONTENT_GAP, 6), pady=PAD_CARD[1])

        # 弱按钮：切换主题（原页眉卡片控件挪入执行区）
        ctk.CTkButton(
            card, text="切换主题", width=90, height=40,
            fg_color="transparent", border_width=1,
            border_color=("#B0B0B0", "#5A5A5A"),
            text_color=("gray40", "gray70"),
            hover_color=("#E8E8E8", "#3A3A3A"),
            command=self.toggle_theme
        ).grid(row=0, column=3, sticky="e", padx=(0, PAD_CARD[0]), pady=PAD_CARD[1])

        card.grid_columnconfigure(1, weight=1)

    # ─── 日志卡片 ─────────────────────────────────────────────────────────

    def _build_log_card(self):
        card, _ = _make_card(self, "运行日志", row=2, column=0)

        self.txt_log = ctk.CTkTextbox(card, wrap="word")
        self.txt_log.grid(row=1, column=0, sticky="nsew",
                          padx=PAD_CARD[0], pady=(0, PAD_CARD[1]))

        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

    # ─── 日志队列 ─────────────────────────────────────────────────────────

    def log(self, msg):
        with self.log_lock:
            self.log_queue.put(msg)

    def _safe_ui(self, callback, *args, **kwargs):
        """将界面更新切回主线程，避免后台线程直接操作 Tk 控件。"""
        self.after(0, lambda: callback(*args, **kwargs))

    def _set_progress(self, value: float):
        """线程安全地更新进度条。"""
        self._safe_ui(self.prog.set, value)
        percent = max(0, min(100, round(value * 100)))
        status = (
            "核对完成"
            if percent == 100
            else ("准备就绪" if percent == 0 else f"正在处理 · {percent}%")
        )
        self._safe_ui(self.status_var.set, status)

    def _set_start_enabled(self, enabled: bool):
        """线程安全地切换开始按钮状态。"""
        self._safe_ui(self.btn_start.configure, state="normal" if enabled else "disabled")

    def _set_stop_enabled(self, enabled: bool):
        """线程安全地切换停止按钮状态。"""
        self._safe_ui(self.btn_stop.configure, state="normal" if enabled else "disabled")

    @staticmethod
    def _is_mapping_selected(value):
        """判断列映射是否已选择有效列。"""
        return value not in (None, "", "(无)")

    @staticmethod
    def _is_supported_file(file_path: Path):
        return file_path.suffix.lower() in {".xlsx", ".xls", ".csv"}

    def _validate_file_path(self, path_value, label):
        path_text = path_value.strip()
        if not path_text:
            return False, f"❌ 请先选择{label}文件"

        path_obj = Path(path_text)
        if not path_obj.exists():
            return False, f"❌ {label}文件不存在: {path_text}"
        if not path_obj.is_file():
            return False, f"❌ {label}路径不是文件: {path_text}"
        if not self._is_supported_file(path_obj):
            return False, f"❌ {label}仅支持 Excel 或 CSV 文件"
        return True, ""

    def _validate_mapping(self, mapping, label, available_columns):
        """校验列映射是否满足当前模式的最低要求。"""
        if not self._is_mapping_selected(mapping.get("date")):
            return False, f"❌ {label}未选择日期列"

        available_set = set(available_columns)
        selected_fields = {}

        def ensure_column(field_key, field_label):
            value = mapping.get(field_key)
            if not self._is_mapping_selected(value):
                return True, ""
            if value not in available_set:
                return False, f"❌ {label}{field_label}不存在于当前文件列中: {value}"
            if value in selected_fields.values():
                return False, f"❌ {label}{field_label}与其他映射重复: {value}"
            selected_fields[field_key] = value
            return True, ""

        ok, message = ensure_column("date", "日期列")
        if not ok:
            return False, message
        ok, message = ensure_column("summary", "摘要列")
        if not ok:
            return False, message

        mode = mapping.get("mode")
        if mode == "debit_credit":
            if not self._is_mapping_selected(mapping.get("debit")):
                return False, f"❌ {label}未选择借方/支出列"
            if not self._is_mapping_selected(mapping.get("credit")):
                return False, f"❌ {label}未选择贷方/收入列"
            ok, message = ensure_column("debit", "借方/支出列")
            if not ok:
                return False, message
            ok, message = ensure_column("credit", "贷方/收入列")
            if not ok:
                return False, message
        elif mode == "single_amount_with_direction":
            if not self._is_mapping_selected(mapping.get("amount")):
                return False, f"❌ {label}未选择金额列"
            if not self._is_mapping_selected(mapping.get("direction")):
                return False, f"❌ {label}未选择方向列"
            ok, message = ensure_column("amount", "金额列")
            if not ok:
                return False, message
            ok, message = ensure_column("direction", "方向列")
            if not ok:
                return False, message
        elif mode == "signed_amount":
            if not self._is_mapping_selected(mapping.get("amount")):
                return False, f"❌ {label}未选择金额列"
            ok, message = ensure_column("amount", "金额列")
            if not ok:
                return False, message

        for optional_key, optional_label in (("voucher", "凭证列"), ("balance", "余额列")):
            if self._is_mapping_selected(mapping.get(optional_key)):
                ok, message = ensure_column(optional_key, optional_label)
                if not ok:
                    return False, message

        for auxiliary_column in mapping.get(
            "auxiliary_text_columns",
            [],
        ):
            if auxiliary_column not in available_set:
                return (
                    False,
                    f"❌ {label}辅助文字列不存在于当前文件中: "
                    f"{auxiliary_column}",
                )

        return True, ""

    @staticmethod
    def _safe_int(value_str: str, default: int = 0) -> int:
        """安全地将字符串转为整数，失败时返回默认值。"""
        try:
            return int(value_str.strip())
        except (ValueError, AttributeError):
            return default

    @staticmethod
    def _safe_float(value_str: str, default: float = 0.0) -> float:
        """安全地将字符串转为浮点数，失败时返回默认值。"""
        try:
            return float(value_str.strip())
        except (ValueError, AttributeError):
            return default

    def _collect_run_state(self):
        """在主线程中一次性快照所有运行参数，避免后台线程直接读取 Tk 对象。"""
        fmt_map = {
            "YYYY-MM-DD": "%Y-%m-%d", "YYYY/MM/DD": "%Y/%m/%d", "YYYYMMDD": "%Y%m%d",
            "DD-MM-YYYY": "%d-%m-%Y", "DD/MM/YYYY": "%d/%m/%Y", "MM/DD/YYYY": "%m/%d/%Y",
            "DD.MM.YYYY": "%d.%m.%Y", "YYYY.MM.DD": "%Y.%m.%d",
        }
        bank_mapping = dict(self.bank_mapping_state)
        journal_mapping = dict(self.journal_mapping_state)
        matcher_values = {
            "performance_materiality": (
                self.performance_materiality_var.get()
            ),
            "clearly_trivial_threshold": (
                self.clearly_trivial_threshold_var.get()
            ),
            "auto_confirm_score": self.auto_confirm_score_var.get(),
            "tolerance_days": self.tolerance_days_var.get(),
            "dfs_window": self.dfs_window_var.get(),
            "dfs_depth": self.dfs_depth_var.get(),
            "greedy_attempts": self.greedy_attempts_var.get(),
            "random_seed": self.random_seed_var.get(),
            "similarity_threshold": self.similarity_threshold_var.get(),
            "similarity_high": self.similarity_high_var.get(),
            "max_candidates": self.max_candidates_var.get(),
            "memory_limit": self.memory_limit_var.get(),
            "batch_min_count": self.batch_min_count_var.get(),
            "allow_mixed_sign": self.allow_mixed_sign_var.get(),
            "allow_zero_match": self.allow_zero_match_var.get(),
            "allow_greedy_fallback": self.allow_greedy_var.get(),
        }
        return {
            "bank_path": self.bank_path.get().strip(),
            "journal_path": self.journal_path.get().strip(),
            "bank_skip": self._safe_int(self.bank_skip.get(), 0),
            "journal_skip": self._safe_int(self.journal_skip.get(), 0),
            "bank_header_rows": self._safe_int(
                self.bank_header_rows.get(),
                1,
            ),
            "journal_header_rows": self._safe_int(
                self.journal_header_rows.get(),
                1,
            ),
            "date_format": fmt_map.get(self.date_fmt_var.get(), "auto"),
            "bank_mapping": bank_mapping,
            "journal_mapping": journal_mapping,
            "config": build_matcher_config(matcher_values),
            "llm_config": self.llm_config,
        }

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.txt_log.insert("end", str(msg) + "\n")
                self.txt_log.see("end")
        except queue.Empty:
            pass
        self.after(100, self.process_log_queue)

    # ─── 文件操作 ─────────────────────────────────────────────────────────

    def browse(self, target):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if path:
            (self.bank_path if target == "bank" else self.journal_path).set(path)
            self.auto_detect(target)

    def auto_detect(self, target):
        path = (self.bank_path if target == "bank" else self.journal_path).get()
        if not path:
            return
        try:
            structure = self.loader.detect_table_structure(path)
        except Exception as exc:
            self.log(f"自动分析表格结构失败: {exc}")
            return
        (self.bank_skip if target == "bank" else self.journal_skip).set(
            str(structure.skiprows)
        )
        (
            self.bank_header_rows
            if target == "bank"
            else self.journal_header_rows
        ).set(str(structure.header_rows))
        label = "银行流水" if target == "bank" else "银行日记账"
        self.log(f"{label}{structure.explanation}")
        self.load_columns(
            path,
            target,
            structure.skiprows,
            structure.header_rows,
            structure.columns,
        )

    def load_columns(
        self,
        path,
        target,
        skiprows=0,
        header_rows=1,
        derived_columns=None,
    ):
        threading.Thread(target=self._load_columns_thread,
                         args=(path, target, skiprows, header_rows, derived_columns), daemon=True).start()

    def _load_columns_thread(
        self,
        path,
        target,
        skiprows,
        header_rows,
        derived_columns,
    ):
        try:
            df = self.loader.load_file(
                path,
                skiprows=int(skiprows),
                header_rows=int(header_rows),
                derived_columns=derived_columns,
            )
            cols = df.columns.tolist()
            self.log(f"已加载列信息: {cols}")
            self.after(
                0,
                lambda: self._store_loaded_columns(target, cols),
            )
        except Exception as e:
            self.log(f"加载列失败: {e}")

    def _store_loaded_columns(self, target, columns):
        if target == "bank":
            self.bank_columns = list(columns)
            self.bank_mapping_state = _auto_mapping_for_columns(
                columns,
                is_bank=True,
            )
        else:
            self.journal_columns = list(columns)
            self.journal_mapping_state = _auto_mapping_for_columns(
                columns,
                is_bank=False,
            )
        if self.bank_columns and self.journal_columns:
            self.mapping_status_var.set("已自动识别，可打开检查")
        else:
            self.mapping_status_var.set("已读取一侧文件列")

    # ─── 主题切换 ─────────────────────────────────────────────────────────

    def toggle_theme(self):
        if self.theme_mode == "system":
            self.theme_mode = "dark"
        elif self.theme_mode == "dark":
            self.theme_mode = "light"
        else:
            self.theme_mode = "system"
        ctk.set_appearance_mode(self.theme_mode)

    # ─── 执行核对 ─────────────────────────────────────────────────────────

    def start(self):
        is_valid, error_msg = validate_config_params(
            self.tolerance_days_var.get(),
            self.dfs_window_var.get(),
            self.dfs_depth_var.get(),
            self.greedy_attempts_var.get(),
            random_seed=self.random_seed_var.get(),
            similarity_threshold=self.similarity_threshold_var.get(),
            similarity_high=self.similarity_high_var.get(),
            max_candidates=self.max_candidates_var.get(),
            memory_limit=self.memory_limit_var.get(),
            bank_skip=self.bank_skip.get(),
            journal_skip=self.journal_skip.get(),
            performance_materiality=(
                self.performance_materiality_var.get()
            ),
            clearly_trivial_threshold=(
                self.clearly_trivial_threshold_var.get()
            ),
            auto_confirm_score=self.auto_confirm_score_var.get(),
        )
        if not is_valid:
            messagebox.showerror("参数错误", error_msg)
            self.log(error_msg)
            return

        try:
            run_state = self._collect_run_state()
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            self.log(str(exc))
            return
        if (
            run_state["bank_header_rows"] < 1
            or run_state["journal_header_rows"] < 1
        ):
            message = "❌ 表头行数必须大于等于1"
            messagebox.showerror("参数错误", message)
            self.log(message)
            return
        for path_value, label in (
            (run_state["bank_path"], "银行流水"),
            (run_state["journal_path"], "日记账"),
        ):
            path_ok, path_error = self._validate_file_path(path_value, label)
            if not path_ok:
                messagebox.showerror("参数错误", path_error)
                self.log(path_error)
                return

        for label, mapping, available_columns in (
            ("银行流水", run_state["bank_mapping"], self.bank_columns),
            ("日记账", run_state["journal_mapping"], self.journal_columns),
        ):
            mapping_ok, mapping_error = self._validate_mapping(mapping, label, available_columns)
            if not mapping_ok:
                messagebox.showerror("参数错误", mapping_error)
                self.log(mapping_error)
                return

        self._set_stop_enabled(True)
        self._set_start_enabled(False)
        threading.Thread(target=self.run_process, args=(run_state,), daemon=True).start()

    def stop_process(self):
        if self.matcher:
            self.matcher.set_stopping(True)
            self.log("正在停止任务...")
            self._set_stop_enabled(False)

    def _confirm_precheck_warnings(
        self,
        report: InputPrecheckReport,
    ) -> bool:
        """在主线程一次展示普通提示，并等待用户选择。"""
        completed = threading.Event()
        decision = {"continue": False}

        def ask_user():
            decision["continue"] = messagebox.askokcancel(
                "输入预检查提示",
                report.warning_message()
                + "\n\n选择“确定”继续核对；选择“取消”返回调整。",
                icon="warning",
                parent=self,
            )
            completed.set()

        self.after(0, ask_user)
        completed.wait()
        return decision["continue"]

    def run_process(self, run_state):
        try:
            self.log("开始处理...")
            output_path = run_reconciliation(
                bank_path=run_state["bank_path"],
                journal_path=run_state["journal_path"],
                bank_mapping=run_state["bank_mapping"],
                journal_mapping=run_state["journal_mapping"],
                matcher_config=run_state["config"],
                llm_config=run_state["llm_config"],
                logger=self.log,
                bank_skiprows=run_state["bank_skip"],
                journal_skiprows=run_state["journal_skip"],
                bank_header_rows=run_state["bank_header_rows"],
                journal_header_rows=run_state["journal_header_rows"],
                date_format=run_state["date_format"],
                progress_callback=self._set_progress,
                matcher_ready=lambda matcher: setattr(
                    self,
                    "matcher",
                    matcher,
                ),
                precheck_warning_callback=self._confirm_precheck_warnings,
            )
            self.log(f"\n{'=' * 50}")
            self.log(f"✅ 核对完成！报告已保存至: {output_path}")
            self.log(f"{'=' * 50}")
        except InputPrecheckBlockedError as exc:
            message = str(exc)
            self.log(f"输入预检查未通过：{message}")
            self.after(
                0,
                lambda: messagebox.showerror(
                    "输入预检查未通过",
                    message,
                    parent=self,
                ),
            )
            self._set_progress(0)
        except InterruptedError as exc:
            self.log(str(exc) or "任务已取消")
            self._set_progress(0)
        except Exception as exc:
            self.log(f"错误: {exc}")
            import traceback
            traceback.print_exc()
        finally:
            self._set_stop_enabled(False)
            self._set_start_enabled(True)
