# -*- coding: utf-8 -*-
"""排版计划：报告排版的前置计算结果（内存对象），供 报告写入.渲染簿计划 单次渲染。

各排版层（通用/复核/展示/紧凑）按与旧代码相同的顺序回放到本结构，
层叠语义与旧路径自然一致。
"""
from dataclasses import dataclass, field

import pandas as pd


@dataclass
class 格样式:
    字体名: str | None = None
    字号: float | None = None
    加粗: bool | None = None
    字色: str | None = None
    下划线: str | None = None
    填充色: str | None = None
    填充类型: str | None = None          # 'solid' 或 'none'
    水平: str | None = None
    垂直: str | None = None
    自动换行: bool | None = None
    上框: tuple | None = None            # (style, color)
    下框: tuple | None = None
    左框: tuple | None = None
    右框: tuple | None = None
    数字格式: str | None = None
    批注: tuple | None = None            # (文本, 作者)
    超链接: str | None = None

    def 属性元组(self):
        return tuple(self.__dict__.values())


def 合并样式(底: 格样式 | None, 盖: 格样式 | None) -> 格样式:
    """盖的非空字段覆盖底；两者皆空返回空样式。"""
    结果 = 格样式()
    for 源 in (底, 盖):
        if 源 is None:
            continue
        for k, v in 源.__dict__.items():
            if v is not None:
                setattr(结果, k, v)
    return 结果


@dataclass
class 列计划:
    宽: float | None = None
    隐藏: bool = False
    基础样式: 格样式 | None = None


@dataclass
class 行计划:
    高: float | None = None
    隐藏: bool = False
    outline: int = 0
    收起: bool = False
    样式: 格样式 | None = None


@dataclass
class 表计划:
    名称: str = ""
    表: pd.DataFrame = field(default_factory=pd.DataFrame)
    表头样式: 格样式 | None = None
    列: list = field(default_factory=list)          # list[列计划]，与 表.columns 对齐
    行: dict = field(default_factory=dict)          # {1 基行号: 行计划}（1=表头）
    覆盖: dict = field(default_factory=dict)        # {(行,列) 1 基: 格样式}
    验证: list = field(default_factory=list)        # DataValidation kwargs（含 sqref）
    条件格式: list = field(default_factory=list)    # [(区域, openpyxl 规则对象)]
    冻结: str | None = None
    筛选: str | None = None
    tab色: str | None = None
    状态: str = "visible"
    网格线: bool = True
    缩放: int | None = None
    左上角: str | None = None
    outline汇总在下: bool | None = None
    打印: dict = field(default_factory=dict)
    前置空列: int = 0
    前部列: list = field(default_factory=list)      # list[列计划]，对应 1..前置空列 的空列

    def 列计划取(self, 下标: int) -> 列计划:
        """取下标（0 基）列计划；不足自动补齐。"""
        while len(self.列) <= 下标:
            self.列.append(列计划())
        return self.列[下标]

    def 行计划取(self, 行号: int) -> 行计划:
        if 行号 not in self.行:
            self.行[行号] = 行计划()
        return self.行[行号]

    def 盖格(self, 行号: int, 列号: int, 样式: 格样式):
        """叠加格覆盖（1 基行列号）。"""
        self.覆盖[(行号, 列号)] = 合并样式(self.覆盖.get((行号, 列号)), 样式)


@dataclass
class 簿计划:
    表: list = field(default_factory=list)          # list[表计划]，顺序即最终表序
    active: int = 0
    自动重算: bool = True

    def 取表(self, 名称: str) -> 表计划 | None:
        for t in self.表:
            if t.名称 == 名称:
                return t
        return None


def 新建簿计划(tables: dict) -> 簿计划:
    """把 {表名: DataFrame} 初始化为裸簿计划（默认样式，未排版）。"""
    计划 = 簿计划()
    for 名称, df in tables.items():
        计划.表.append(表计划(名称=名称, 表=df))
    return 计划


# ═══════════════════════════════════════════════════════════════════════════
# 通用层回放（翻译自 make_excel._apply_styles 与 reporter 的三个后处理 +
# _apply_report_presentation；逐条对照源逻辑，同序回放保证层叠语义一致）
# ═══════════════════════════════════════════════════════════════════════════

import math as _math

from openpyxl.utils import get_column_letter as _字母


def 应用底层主题(tp: 表计划, theme: str = "deep-navy"):
    """翻译 make_excel._apply_styles（start_col=2，即前置空列=1）。"""
    from make_excel import (
        A_COL_WIDTH, FMT, ROW_HEIGHT, SUM_KEYWORDS, THEMES,
        _estimate_column_width, _infer_column_type,
    )

    t = THEMES[theme]
    df = tp.表
    tp.前置空列 = 1
    ncols = len(df.columns)
    nrows = 1 + len(df)
    for r in range(1, nrows + 1):
        tp.行计划取(r).高 = ROW_HEIGHT
    # A 列永远无数据（_write_data 从 B 列写起）→ 走 else 分支：宽 2
    tp.前部列 = [列计划(宽=A_COL_WIDTH)]

    # 表头：主题表头样式 + 首行边框（TOP_BORDER）
    tp.表头样式 = 格样式(
        字体名="Arial", 字号=11, 加粗=True, 字色=t["header_font_color"],
        填充色=t["header_fill"], 填充类型="solid",
        水平="right", 垂直="center",
        上框=("medium", "000000"), 下框=("dashed", "808080"),
    )

    # 列类型推断 + 列宽估算 + 数据格基础样式
    for ci in range(ncols):
        col_type = _infer_column_type(str(df.columns[ci]), df.iloc[:, ci])
        cp = tp.列计划取(ci)
        cp.宽 = _estimate_column_width(str(df.columns[ci]), df.iloc[:, ci])
        is_num = col_type in ("money", "number", "pct")
        nf = FMT.get(col_type)
        cp.基础样式 = 格样式(
            字体名="Arial", 字号=11,
            字色=t["data_font_blue"] if is_num else t["data_font_black"],
            水平="right" if is_num else "left", 垂直="center",
            数字格式=nf or None,
        )
        if col_type == "date":
            # 旧逻辑：date 列里的原始数字显式回落 General，不当日期序号
            for ri, v in enumerate(df.iloc[:, ci].tolist()):
                if isinstance(v, (int, float)) and not isinstance(v, bool) \
                        and not (isinstance(v, float) and _math.isnan(v)):
                    tp.盖格(ri + 2, ci + 2, 格样式(数字格式="General"))

    # 合计行检测（末行首数据列含合计关键词）
    if len(df) > 0 and ncols > 0:
        last_val = df.iloc[-1, 0]
        if last_val is not None and any(kw in str(last_val) for kw in SUM_KEYWORDS):
            rp = tp.行计划取(nrows)
            rp.样式 = 合并样式(rp.样式, 格样式(
                字体名="Arial", 字号=11, 加粗=t["sum_font_bold"],
                字色=t["sum_font_color"], 填充色=t["sum_fill"], 填充类型="solid",
            ))

    # 边框：首行 TOP（已并入表头样式）、末行 BOTTOM、中间 MID
    中虚 = ("dashed", "808080")
    for r in range(2, nrows + 1):
        if r == nrows:
            边 = 格样式(上框=中虚, 下框=("medium", "000000"))
        else:
            边 = 格样式(上框=中虚, 下框=中虚)
        rp = tp.行计划取(r)
        rp.样式 = 合并样式(rp.样式, 边)

    # 右侧空白列（宽 3）
    tp.列计划取(ncols).宽 = 3

    tp.网格线 = False
    tp.冻结 = "A2"
    if ncols > 0:
        tp.筛选 = f"B1:{_字母(1 + ncols)}{max(1, nrows)}"


def 应用明细着色(tp: 表计划):
    """翻译 reporter._postprocess_details；_低置信度标记列直接 drop（等价于
    旧代码先着色后删列——着色覆盖全部数据列，删列后效果不变）。"""
    df = tp.表
    cols = list(df.columns)
    if "_低置信度标记" in cols:
        idx = cols.index("_低置信度标记")
        for ri, v in enumerate(df.iloc[:, idx].tolist()):
            if v is True or str(v).lower() == "true":
                rp = tp.行计划取(ri + 2)
                rp.样式 = 合并样式(rp.样式, 格样式(
                    字体名="", 字号="", 加粗=True, 字色="FF0000",
                    填充色="FFFF00", 填充类型="solid",
                ))
        tp.表 = df.drop(columns=["_低置信度标记"])
        if len(tp.列) > idx:
            tp.列.pop(idx)
        df = tp.表
        cols = list(df.columns)

    risk_idx = cols.index("风险等级") if "风险等级" in cols else None
    风险色 = {"低风险": "FFF2CC", "中风险": "FCE4D6", "高风险": "FFC7CE", "范围未知": "F4CCCC"}
    diff_indices = [i for i, h in enumerate(cols)
                    if "差额" in str(h) or "diff" in str(h).lower()]
    for ri in range(len(df)):
        for di in diff_indices:
            v = df.iloc[ri, di]
            try:
                f = float(v) if v is not None else 0
            except (ValueError, TypeError):
                continue
            if f != 0:
                risk = df.iloc[ri, risk_idx] if risk_idx is not None else "低风险"
                tp.盖格(ri + 2, di + 2, 格样式(
                    填充色=风险色.get(str(risk), 风险色["低风险"]), 填充类型="solid"))


def 应用差额着色(tp: 表计划):
    """翻译 reporter._postprocess_diff_columns：非零差额统一 FFF2CC。"""
    df = tp.表
    cols = list(df.columns)
    for ci, h in enumerate(cols):
        if "差额" not in str(h) and "diff" not in str(h).lower():
            continue
        for ri in range(len(df)):
            v = df.iloc[ri, ci]
            try:
                f = float(v) if v is not None else 0
            except (ValueError, TypeError):
                continue
            if f != 0:
                tp.盖格(ri + 2, ci + 2, 格样式(填充色="FFF2CC", 填充类型="solid"))


def 应用期初警告(tp: 表计划, has_warning: bool):
    """翻译 reporter._postprocess_summary（两次调用幂等，计划层同序回放两次）。"""
    if not has_warning:
        return
    df = tp.表
    cols = list(df.columns)
    if "项目" not in cols:
        return
    item_idx = cols.index("项目")
    for ri in range(len(df)):
        if "期初余额" in str(df.iloc[ri, item_idx] or ""):
            rp = tp.行计划取(ri + 2)
            rp.样式 = 合并样式(rp.样式, 格样式(
                字体名="", 字号=14, 加粗=True, 字色="FF0000",
                填充色="FFFF00", 填充类型="solid",
            ))


_技术表 = {"运行参数", "大模型辅助明细", "解析异常明细"}
_技术列 = {
    "候选ID", "阶段", "最终状态", "处理原因", "综合可信度",
    "金额分", "日期分", "文字分", "结构分", "银_日期", "账_日期",
    "文字支持", "文字冲突", "大模型判断", "差异池ID", "是否使用大模型",
    "纳入风险池", "匹配候选ID", "银行收入", "银行支出", "银行净额",
    "日记账收入", "日记账支出", "日记账净额", "组金额", "收入差额",
    "支出差额",
}
_疑点隐藏列 = {
    "银行原文件行号", "日记账原文件行号", "银行日期", "日记账日期",
    "银行金额", "日记账金额", "调整凭证号", "责任人", "处理日期", "匹配ID",
}
_金额关键词 = ("金额", "合计", "收入", "支出", "净额", "差额", "余额", "组金额")
_宽列 = {"系统结论", "判断依据", "业务分组依据", "其他可能对应", "建议动作",
         "摘要", "银行摘要", "日记账摘要", "说明"}
_百分比表头 = {
    "逐笔精确匹配率", "组级勾稽率", "自动完成率",
    "银行已核对笔数覆盖率", "银行已核对金额覆盖率",
    "序时账已核对笔数覆盖率", "序时账已核对金额覆盖率",
}


def _中文宽(text: str) -> int:
    return sum(2 if ord(c) > 127 else 1 for c in text)


def 应用通用排版(簿: 簿计划):
    """翻译 reporter._apply_report_presentation。"""
    for tp in 簿.表:
        df = tp.表
        cols = list(df.columns)
        ncols = len(cols)
        nrows = 1 + len(df)
        偏 = tp.前置空列
        if tp.名称 in _技术表:
            tp.状态 = "hidden"
        first = 偏 + 1 if ncols > 0 else 1
        if first > 1:
            for k in range(first - 1):
                tp.前部列[k].隐藏 = True
        tp.冻结 = f"{_字母(first)}2"
        if nrows >= 1 and ncols >= 1:
            tp.筛选 = f"{_字母(first)}1:{_字母(偏 + ncols)}{nrows}"

        # sheet[1] 遍历覆盖 A 列：值 None → 宽 14；数据行落顶端对齐空格
        if first > 1:
            for k in range(first - 1):
                tp.前部列[k].宽 = 14
                if len(df) > 0:
                    tp.前部列[k].基础样式 = 格样式(垂直="top", 自动换行=True)

        for j, col_name in enumerate(cols):
            cp = tp.列计划取(j)
            if col_name in _技术列:
                cp.隐藏 = True
            if tp.名称 == "疑点事项" and col_name in _疑点隐藏列:
                cp.隐藏 = True
            width = 28 if col_name in _宽列 else 14
            cp.宽 = width
            if len(df) == 0:
                continue
            base = cp.基础样式 or 格样式()
            base.水平 = ""          # 旧代码整体替换 Alignment，horizontal 被清空
            base.垂直 = "top"
            base.自动换行 = True
            if any(k in str(col_name) for k in _金额关键词):
                base.数字格式 = "#,##0.00"
            cp.基础样式 = base
            if cp.隐藏:
                continue
            for ri, v in enumerate(df.iloc[:, j].tolist()):
                if isinstance(v, str):
                    lines = sum(
                        max(1, (_中文宽(line) + width - 1) // width)
                        for line in v.split("\n")
                    )
                    rp = tp.行计划取(ri + 2)
                    rp.高 = min(409, max(rp.高 or 20, lines * 16 + 4))

        if tp.名称 == "核对结论":
            item_idx = cols.index("项目") if "项目" in cols else 0
            value_idx = cols.index("数值") if "数值" in cols else item_idx + 1
            tp.列计划取(item_idx).宽 = 24
            if value_idx < ncols:
                tp.列计划取(value_idx).宽 = 58
            for ri in range(len(df)):
                if df.iloc[ri, item_idx] in _百分比表头:
                    tp.盖格(ri + 2, 偏 + value_idx + 1, 格样式(数字格式="0.00%"))
                value = df.iloc[ri, value_idx] if value_idx < ncols else None
                text_width = _中文宽(str(value or ""))
                tp.行计划取(ri + 2).高 = min(
                    409,
                    max(32 if len(str(value or "")) > 20 else 20,
                        ((text_width + 51) // 52) * 16 + 4),
                )

        if tp.名称 == "输入检查":
            widths = {"检查项目": 20, "银行流水结果": 20, "银行日记账结果": 20,
                      "双方比较结果": 24, "状态": 12, "说明": 52}
            for header, w in widths.items():
                if header in cols:
                    tp.列计划取(cols.index(header)).宽 = w
            for ri in range(len(df)):
                longest = max(
                    (len(str(v or "")) for v in df.iloc[ri].tolist()),
                    default=0,
                )
                tp.行计划取(ri + 2).高 = 36 if longest > 18 else 22

    # 疑点事项验证 + 核对结论公式（跨表部分）
    疑点 = 簿.取表("疑点事项")
    if 疑点 is not None:
        cols = list(疑点.表.columns)
        status_idx = cols.index("后续状态") if "后续状态" in cols else None
        if status_idx is not None:
            letter = _字母(疑点.前置空列 + status_idx + 1)
            疑点.验证.append({
                "type": "list",
                "formula1": '"已关注,已处理,无需处理"',
                "allow_blank": True,
                "sqref": f"{letter}2:{letter}{max(2, len(疑点.表) + 1)}",
            })
        结论 = 簿.取表("核对结论")
        if 结论 is not None and status_idx is not None and "系统结论" in cols:
            item_idx = cols.index("系统结论")
            last_row = max(2, len(疑点.表) + 1)
            status_letter = _字母(疑点.前置空列 + status_idx + 1)
            item_letter = _字母(疑点.前置空列 + item_idx + 1)
            status_range = f"'疑点事项'!${status_letter}$2:${status_letter}${last_row}"
            item_range = f"'疑点事项'!${item_letter}$2:${item_letter}${last_row}"
            formulas = {
                "待处理事项数": f'=COUNTIFS({item_range},"<>",{status_range},"")',
                "已关注事项数": f'=COUNTIF({status_range},"已关注")',
                "已处理事项数": f'=COUNTIF({status_range},"已处理")',
                "无需处理事项数": f'=COUNTIF({status_range},"无需处理")',
            }
            结论cols = list(结论.表.columns)
            if "项目" in 结论cols:
                si = 结论cols.index("项目")
                vi = 结论cols.index("数值") if "数值" in 结论cols else si + 1
                for ri in range(len(结论.表)):
                    item = 结论.表.iloc[ri, si]
                    if item in formulas:
                        结论.表.iat[ri, vi] = formulas[item]
                簿.自动重算 = True
