# -*- coding: utf-8 -*-
"""排版共享：排版路径的纯逻辑常量与工具（零 openpyxl）。

这些内容原本散在 make_excel.py / 报告展示.py / 紧凑报告.py（旧排版遗留模块，
均顶层依赖 openpyxl）。生产单次写入路径只从这里取用，保证生成报告不加载
openpyxl；旧模块保留自己的副本继续服务回退路径。
"""
import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd


# ═══════════════════════════════════════════════════════════════════════════
# 列字母
# ═══════════════════════════════════════════════════════════════════════════

_列字母缓存 = {}


def 字母(col: int) -> str:
    """1 基列号 → Excel 字母（带缓存）。"""
    s = _列字母缓存.get(col)
    if s is None:
        n, s = col, ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        _列字母缓存[col] = s
    return s


# ═══════════════════════════════════════════════════════════════════════════
# 条件格式规则（自有结构，替代 openpyxl FormulaRule）
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class 条件规则:
    """表达式条件格式：公式命中时套填充色/字色。"""
    公式: list = field(default_factory=list)
    填充色: str | None = None
    字色: str | None = None


# ═══════════════════════════════════════════════════════════════════════════
# 阅读表与组成金额列（原 报告展示.py）
# ═══════════════════════════════════════════════════════════════════════════

阅读表 = ("核对结论", "非工作日交易", "月度核对", "逐笔核对", "整组核对",
          "人工全查", "人工抽样", "未对应记录")
组成金额列 = ("银行收入", "银行支出", "序时账收入", "序时账支出")


# ═══════════════════════════════════════════════════════════════════════════
# 紧凑层纯函数（原 紧凑报告.py）
# ═══════════════════════════════════════════════════════════════════════════

def 短结果(expression):
    """原 紧凑报告._short_result：详细状态缩短为显示文字。"""
    replacements = (
        ("所选对应重复使用原记录", "重复占用，请重选"),
        ("原对应已否定，记录恢复未对应", "已取消对应"),
        ("原对应被人工选择替代，记录未对应", "原对应已被替代"),
        ("已人工核对，保留未对应", "已查看，仍未对应"),
        ("已确认所选对应；原始差额及期间保留", "已确认，差额保留"),
        ("部分原记录已由其他事项建立对应，其余仍未对应", "部分记录已对应"),
        ("原记录已由其他事项建立对应", "已由其他事项对应"),
        ("程序已确认对应", "程序已对上"),
        ("待人工核对", "待核对"),
        ("留存备查，未作人工确认", "留存查看"),
        ("选择无效", "选择无效，请重选"),
    )
    for old, new in replacements:
        expression = f'SUBSTITUTE({expression},"{old}","{new}")'
    return "=" + expression


def 入选原因(item, bank, journal, income_diff, expense_diff):
    """原 紧凑报告._plain_reason。"""
    parts = []
    if not bank:
        parts.append("账上有记录，尚未找到银行对应")
    elif not journal:
        parts.append("银行有记录，尚未找到账上对应")
    elif income_diff or expense_diff:
        parts.append("找到可能对应，但收支金额有差额")
    else:
        parts.append("金额相同，对应关系仍需确认")
    if item.get("跨期"):
        parts.append("双方记录跨月")
    if item.get("风险等级") == "范围未知":
        parts.append("资料或余额有缺口")
    reason = str(item.get("判断依据") or "")
    if "冲突" in reason:
        parts.append("关键业务信息有冲突")
    if "批次边界" in reason:
        parts.append("同一批业务的范围不清楚")
    if "多个" in reason or "歧义" in reason:
        parts.append("有多种可能对应")
    return "；".join(parts)


def 核对动作(item, bank, journal, income_diff, expense_diff):
    """把待办事项翻译成可以直接执行的核对动作。"""
    if not bank:
        return "查银行流水有无这笔账；找到后核对日期、金额和对方，选黄色格；找不到则备注原因"
    if not journal:
        return "查序时账有无这笔银行收付；找到后核对日期、金额和对方，选黄色格；找不到则备注原因"
    reason = str(item.get("判断依据") or "")
    if "冲突" in reason:
        return "核对双方对方名称、账号等冲突信息；确属同一笔才选确认对应，否则选否定对应"
    if "多个" in reason or "歧义" in reason:
        action = "比较完整明细中的其他对应方案，核实原表记录后在黄色格选正确关系"
        return action + "，并在备注写明差额原因" if income_diff or expense_diff else action
    if income_diff or expense_diff:
        return "核对是否同一笔并查明收支差额；在黄色格选结果，备注写明差额原因"
    if len(bank) > 1 or len(journal) > 1:
        return "逐笔核对双方完整组成、金额及业务范围；一致才在黄色格选确认对应"
    return "核对日期、金额、对方名称和业务说明；一致选黄色格确认对应，否则选否定对应"


def 记录描述(records):
    """原 紧凑报告._description。"""
    if not records:
        return "尚未找到对应记录"
    summaries = list(dict.fromkeys(str(row.get("摘要") or "") for row in records if row.get("摘要")))
    parties = list(dict.fromkeys(
        part for row in records
        for part in str(row.get("对方及凭证") or "").split("；")
        if part and part not in summaries
    ))
    days = sorted({str(row["日期"])[:10] for row in records if row.get("日期")})
    day = days[0] if len(days) == 1 else days[0] + "至" + days[-1] if days else ""
    source_rows = list(dict.fromkeys(
        match.group(1)
        for row in records
        if (match := re.search(r"第(\d+)行", str(row.get("原始出处") or "")))
    ))
    source_hint = "原表第" + "、".join(source_rows[:3]) + "行" if source_rows else ""
    if len(source_rows) > 3:
        source_hint += "等"
    content = "；".join(summaries[:2] + parties[:2])
    if len(summaries) > 2 or len(parties) > 2:
        content += "等，完整组成见明细"
    return " ".join(part for part in (day, f"{len(records)}笔", source_hint, content) if part)


def 记录金额(records):
    """原 紧凑报告._amounts：返回 (收入合计, 支出合计)，负数归入支出。"""
    values = [Decimal(str(row.get("金额") or 0)) for row in records]
    return (float(sum((max(x, Decimal(0)) for x in values), Decimal(0))),
            float(sum((max(-x, Decimal(0)) for x in values), Decimal(0))))


# ═══════════════════════════════════════════════════════════════════════════
# 底层主题常量与列推断（原 make_excel.py 纯逻辑部分）
# ═══════════════════════════════════════════════════════════════════════════

SUM_KEYWORDS = ("合计", "小计", "总计", "汇总", "sum", "total", "subtotal", "小計", "合計")

FMT = {
    "money": "#,##0.00",
    "number": "#,##0",
    "pct": "0.00%",
    "date": "yyyy/mm/dd",
    "text": "@",
    "unknown": "@",
}

ROW_HEIGHT = 18
A_COL_WIDTH = 2
MIN_COL_WIDTH = 10
MAX_COL_WIDTH = 50

THEMES = {
    "default": {
        "label": "水蓝",
        "header_fill": "4472C4",
        "header_font_color": "FFFFFF",
        "sum_fill": "D9E2F3",
        "sum_font_bold": True,
        "sum_font_color": "000000",
        "data_font_blue": "0000FF",
        "data_font_black": "000000",
        "desc": "经典水蓝（原摩根系默认）",
    },
    "deep-navy": {
        "label": "深海蓝",
        "header_fill": "1F4E79",
        "header_font_color": "FFFFFF",
        "sum_fill": "D6E4F0",
        "sum_font_bold": True,
        "sum_font_color": "000000",
        "data_font_blue": "0000FF",
        "data_font_black": "000000",
        "desc": "深海蓝表头，更沉稳专业",
    },
    "jade": {
        "label": "墨玉绿",
        "header_fill": "375623",
        "header_font_color": "FFFFFF",
        "sum_fill": "E2EFDA",
        "sum_font_bold": True,
        "sum_font_color": "000000",
        "data_font_blue": "0000FF",
        "data_font_black": "000000",
        "desc": "墨绿表头，清爽自然",
    },
    "slate": {
        "label": "陨石灰蓝",
        "header_fill": "404040",
        "header_font_color": "FFFFFF",
        "sum_fill": "D9D9D9",
        "sum_font_bold": True,
        "sum_font_color": "000000",
        "data_font_blue": "0000FF",
        "data_font_black": "000000",
        "desc": "深灰表头 + 蓝色色阶，现代极简",
    },
    "burgundy": {
        "label": "勃艮第红",
        "header_fill": "843C0C",
        "header_font_color": "FFFFFF",
        "sum_fill": "FCE4D6",
        "sum_font_bold": True,
        "sum_font_color": "000000",
        "data_font_blue": "0000FF",
        "data_font_black": "000000",
        "desc": "酒红表头，暖色调高级感",
    },
    "coral": {
        "label": "珊瑚橙",
        "header_fill": "D84B4B",
        "header_font_color": "FFFFFF",
        "sum_fill": "FDE8E8",
        "sum_font_bold": True,
        "sum_font_color": "000000",
        "data_font_blue": "0000FF",
        "data_font_black": "000000",
        "desc": "珊瑚红表头，活泼明亮",
    },
}


def 列类型推断(name: str, series: pd.Series) -> str:
    """原 make_excel._infer_column_type：列名关键词 + 值特征推断列类型。"""
    n = name.lower()
    if re.search(r"(率$|占比|百分比|增长率|利润率|达成率|rate$|ratio$|margin|pct|percent)", n):
        return "pct"
    if re.search(r"(日期|时间|年月|期间|年份|月份|date|time|year|month)", n):
        return "date"
    if re.search(r"(编号|id|单号|编码|电话|手机|备注|说明|名称|地址|描述|号码|订单号|负责人|姓名|联系人|部门|岗位|phone|email|code|desc|address)", n):
        return "text"
    if re.search(r"(金额|余额|价格|收入|成本|费用|毛利额|净利|合计|总额|售价|单价|预算|支出|毛利(?!率)|返点|佣金|折扣|补贴|周转|本金|amount|balance|price|cost|revenue|total|budget|expense|commission|discount|rebate|profit|spending)", n):
        return "money"

    clean = series.dropna()
    if len(clean) == 0:
        return "unknown"
    try:
        if pd.api.types.is_datetime64_any_dtype(clean):
            return "date"
        if pd.api.types.is_numeric_dtype(clean):
            sample = clean.head(20)
            if pd.api.types.is_float_dtype(sample):
                if (sample >= 0).all() and (sample <= 1).all() and sample.nunique() > 2:
                    return "pct"
                return "money"
            return "number"
        str_sample = clean.astype(str).head(20)
        if str_sample.str.match(r"^\d{12,}$").any():
            return "text"
    except Exception:
        pass
    return "unknown"


def _字符宽(text) -> int:
    """中文字符计 2，拉丁/数字计 1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))


def 列宽估算(header: str, series: pd.Series) -> int:
    """原 make_excel._estimate_column_width：表头和数据最大字符宽 + padding。

    超过 5000 行的表只取样前 500 个非空值：列宽是视觉近似（用户口径
    "看起来差不多"），30 万行全量逐字计算没有意义且占排版耗时的近一成；
    小表全量扫描，保证金样逐格一致。
    """
    max_w = _字符宽(header)
    值列 = series.dropna()
    if len(值列) > 5000:
        值列 = 值列.head(500)
    for v in 值列:
        try:
            w = _字符宽(str(v))
            if w > max_w:
                max_w = w
        except Exception:
            pass
    return max(MIN_COL_WIDTH, min(max_w + 2, MAX_COL_WIDTH))


# 兼容别名：排版计划沿用 make_excel 时代的原名导入
_infer_column_type = 列类型推断
_estimate_column_width = 列宽估算


@contextmanager
def atomic_output_path(output_path):
    """先在同盘临时位置完成写入，Windows 改名时拒绝覆盖同名文件。"""
    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="核对暂存_", dir=destination.parent) as directory:
        temporary = Path(directory) / destination.name
        yield temporary
        try:
            temporary.rename(destination)
        except FileExistsError:
            raise FileExistsError(
                f"输出文件已存在，未覆盖：{destination}\n"
                "请先关闭并移走（或改名）旧文件后再试。"
            ) from None
