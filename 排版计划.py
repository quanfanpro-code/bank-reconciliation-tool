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
    条件格式: list = field(default_factory=list)    # [(区域, 条件规则)]
    冻结: str | None = None
    筛选: str | None = None
    tab色: str | None = None
    状态: str = "visible"
    网格线: bool = True
    缩放: int | None = None
    左上角: str | None = None
    outline汇总在下: bool | None = None
    显示outline符号: bool | None = None
    表头含前列: bool = False                 # True 时表头样式也套到前置空列
    打印: dict = field(default_factory=dict)
    前置空列: int = 0
    前部列: list = field(default_factory=list)      # list[列计划]，对应 1..前置空列 的空列
    默认行高: float | None = None
    格变换: object = None                           # callable(格样式, 值, 列标签) -> 格样式，数据格合并后调用
    后覆盖: dict = field(default_factory=dict)      # {(行,列)}，格变换之后再合并的格覆盖
    空缺格: set = field(default_factory=set)        # {(行,列)}，旧路径从不存在的格，渲染时不落格
    物化空格: bool = False                          # True 时 None 值格也走格变换并落格（紧凑层新建表）

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

from 排版共享 import 字母 as _字母, 条件规则

# 大表快速通道阈值：超过此行数的表不再逐格刷样式，改用列默认样式（用户口径：
# 排版是锦上添花，逐格一致只对小表保留）。金样账户最大表 3531 行（修竹四分
# 复核候选组成），阈值 5000 保证三份金样全部走逐格路径、继续作回归基准
_快速通道行数 = 5000


def 应用底层主题(tp: 表计划, theme: str = "deep-navy"):
    """翻译 make_excel._apply_styles（start_col=2，即前置空列=1）。"""
    from 排版共享 import (
        A_COL_WIDTH, FMT, ROW_HEIGHT, SUM_KEYWORDS, THEMES,
        _estimate_column_width, _infer_column_type,
    )

    t = THEMES[theme]
    df = tp.表
    tp.前置空列 = 1
    ncols = len(df.columns)
    nrows = 1 + len(df)
    大表 = len(df) > _快速通道行数
    if 大表:
        # 快速通道：逐行行高/行边框是几十万次样式写入的大头，抹为表级默认行高
        tp.默认行高 = ROW_HEIGHT
    else:
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
            # 旧代码文本列 Font(name='Arial', size=11) 不带字色，仅数字列蓝色
            字色=t["data_font_blue"] if is_num else None,
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
    # （快速通道大表跳过：逐行虚线边框在留痕大表上无阅读价值，且是逐格写入大头）
    if not 大表:
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
    risk_vals = df.iloc[:, risk_idx].tolist() if risk_idx is not None else None
    for di in diff_indices:
        vals = pd.to_numeric(df.iloc[:, di], errors="coerce").fillna(0)
        for ri in vals[vals != 0].index:
            risk = risk_vals[ri] if risk_vals is not None else "低风险"
            tp.盖格(ri + 2, di + 2, 格样式(
                填充色=风险色.get(str(risk), 风险色["低风险"]), 填充类型="solid"))


def 应用差额着色(tp: 表计划):
    """翻译 reporter._postprocess_diff_columns：非零差额统一 FFF2CC。"""
    df = tp.表
    cols = list(df.columns)
    for ci, h in enumerate(cols):
        if "差额" not in str(h) and "diff" not in str(h).lower():
            continue
        vals = pd.to_numeric(df.iloc[:, ci], errors="coerce").fillna(0)
        for ri in vals[vals != 0].index:
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


# ═══════════════════════════════════════════════════════════════════════════
# 复核层回放（翻译自 复核报表.apply_review_presentation / _summary_formulas /
# _format_sheets；逐条对照源逻辑）
# ═══════════════════════════════════════════════════════════════════════════

from collections import defaultdict as _defaultdict


def _列定位(tp: 表计划) -> dict:
    """报告列识别.identify_columns 的计划层等价：表头 → 绝对列号。"""
    定位 = {}
    for i, c in enumerate(tp.表.columns):
        label = str(c).strip() if c is not None else ""
        if not label:
            continue
        if label in 定位:
            raise ValueError(f"工作表“{tp.名称}”存在重复列名“{label}”，请区分列名后再操作")
        定位[label] = tp.前置空列 + i + 1
    return 定位


def _写值(tp: 表计划, row: int, col: int, value):
    j = col - tp.前置空列 - 1
    try:
        tp.表.iat[row - 2, j] = value
    except (TypeError, ValueError):
        # 列 dtype 不兼容（如 float 列写文本）→ 转 object 再写
        列名 = tp.表.columns[j]
        tp.表[列名] = tp.表[列名].astype(object)
        tp.表.iat[row - 2, j] = value


def _读值(tp: 表计划, row: int, col: int):
    """读计划值；NaN/NaT/空串归一为 None（旧路径保存重载后三者即 None，语义对齐）。"""
    v = tp.表.iat[row - 2, col - tp.前置空列 - 1]
    if v is None or v is pd.NaT or v == "":
        return None
    if isinstance(v, float) and _math.isnan(v):
        return None
    return v


def 应用复核层(簿: 簿计划):
    """翻译 复核报表.apply_review_presentation（末尾的 format_display_sheets
    由 应用展示层 紧接回放，保持同样的调用次序）。"""
    if 簿.取表("复核事项索引") is None:
        return
    sheets = {}
    for n in ("复核事项索引", "复核候选选项", "复核候选组成", "核对明细",
              "月度核对", "人工全查", "人工抽样"):
        t = 簿.取表(n)
        if t is None:
            raise KeyError(n)
        sheets[n] = t
    headers = {n: _列定位(t) for n, t in sheets.items()}

    def 行数(n):
        return len(sheets[n].表) + 1

    def cell(n, label, row):
        return f"{_字母(headers[n][label])}{row}"

    def rng(n, label):
        letter = _字母(headers[n][label])
        return f"'{n}'!${letter}$2:${letter}${max(2, 行数(n))}"

    def put(n, label, row, formula):
        _写值(sheets[n], row, headers[n][label], formula)

    def find(n, label, key, key_label="事项编号", default='""'):
        return f"IFERROR(INDEX({rng(n,label)},MATCH({key},{rng(n,key_label)},0)),{default})"

    ix, op, mem, detail = "复核事项索引", "复核候选选项", "复核候选组成", "核对明细"
    index_rows = {str(_读值(sheets[ix], r, headers[ix]["事项编号"])): r
                  for r in range(2, 行数(ix) + 1)}
    option_rows = _defaultdict(list)
    for row in range(2, 行数(op) + 1):
        item_id = str(_读值(sheets[op], row, headers[op]["事项编号"]))
        option_rows[item_id].append(row)

    # 稳定编号查找人工输入
    for item_id, row in index_rows.items():
        key = cell(ix, "事项编号", row)
        mode = _读值(sheets[ix], row, headers[ix]["核对方式"])
        raw_choice = find(mode, "人工核对结果", key, "填报编号") \
            if mode in ("人工全查", "人工抽样") else '""'
        choice = f'IF({raw_choice}="","",{raw_choice})'
        put(ix, "人工核对结果", row, "=" + choice)

    for mode in ("人工全查", "人工抽样"):
        tp = sheets[mode]
        for row in range(2, 行数(mode) + 1):
            item_id = _读值(tp, row, headers[mode]["事项编号"])
            if not item_id:
                continue
            if _读值(tp, row, headers[mode]["行别"]) != "核对事项":
                put(mode, "人工核对结果", row,
                    "=" + find(ix, "人工核对结果", cell(mode, "事项编号", row)))
                put(mode, "当前核对结论", row,
                    "=" + find(ix, "当前核对结论", cell(mode, "事项编号", row)))
                continue
            choices = [_读值(sheets[op], r, headers[op]["选择文字"])
                       for r in option_rows.get(str(item_id), [])]
            choices.append("否定对应" if choices else "保留未对应")
            tp.验证.append({
                "type": "list",
                "formula1": '"' + ",".join(choices) + '"',
                "allow_blank": True,
                "errorTitle": "请选择本事项的结果",
                "error": "请选择下拉列表中的对应结果。",
                "showErrorMessage": True,
                "errorStyle": "stop",
                "sqref": cell(mode, "人工核对结果", row),
            })
            put(mode, "当前核对结论", row,
                "=" + find(ix, "当前核对结论", cell(mode, "事项编号", row)))

    for row in range(2, 行数(op) + 1):
        key = cell(op, "事项编号", row)
        choice = find(ix, "人工核对结果", key)
        put(op, "人工选中", row, f'=--({choice}={cell(op,"选择文字",row)})')

    for row in range(2, 行数(mem) + 1):
        opt = cell(mem, "选项编号", row)
        key = cell(mem, "原始记录号", row)
        put(mem, "人工占用", row, "=" + find(op, "人工选中", opt, "选项编号", "0"))
        put(mem, "占用次数", row,
            f'=SUMIF({rng(mem,"原始记录号")},{key},{rng(mem,"人工占用")})')

    for row in range(2, 行数(op) + 1):
        opt = cell(op, "选项编号", row)
        manual = cell(op, "人工选中", row)
        choice = find(ix, "人工核对结果", cell(op, "事项编号", row))
        put(op, "选择冲突", row,
            f'=COUNTIFS({rng(mem,"选项编号")},{opt},{rng(mem,"占用次数")},">1")')
        put(op, "有效对应", row,
            f'=IF({manual}=1,--({cell(op,"选择冲突",row)}=0),'
            f'IF(AND({cell(op,"自动确认",row)}=1,{choice}=""),'
            f'--(SUMIF({rng(mem,"选项编号")},{opt},{rng(mem,"占用次数")})=0),0))')

    for row in range(2, 行数(mem) + 1):
        opt = cell(mem, "选项编号", row)
        put(mem, "实际生效", row, "=" + find(op, "有效对应", opt, "选项编号", "0"))
        put(mem, "生效记录号", row,
            f'=IF({cell(mem,"实际生效",row)}=1,{cell(mem,"原始记录号",row)},"")')
        put(mem, "显示结果", row, "=" + find(op, "显示结果", opt, "选项编号"))
        put(mem, "已对平", row,
            f'=IF({cell(mem,"实际生效",row)}=1,'
            f'--({find(op,"差异金额",opt,"选项编号","1")}=0),0)')

    for item_id, row in index_rows.items():
        key = cell(ix, "事项编号", row)
        choice = cell(ix, "人工核对结果", row)
        selected = f'SUMIFS({rng(op,"人工选中")},{rng(op,"事项编号")},{key})'
        conflict = f'SUMIFS({rng(op,"选择冲突")},{rng(op,"事项编号")},{key},{rng(op,"人工选中")},1)'
        effective = f'SUMIFS({rng(op,"有效对应")},{rng(op,"事项编号")},{key})'
        mode = _读值(sheets[ix], row, headers[ix]["核对方式"])
        default = ("待人工核对" if mode in ("人工全查", "人工抽样")
                   else "原对应被人工选择替代，记录未对应" if mode == "自动确认"
                   else "留存备查，未作人工确认")
        formula = (f'=IF({choice}="否定对应","原对应已否定，记录恢复未对应",'
                   f'IF({choice}="保留未对应","已人工核对，保留未对应",'
                   f'IF({selected}>0,IF({conflict}>0,"所选对应重复使用原记录",'
                   f'"已确认所选对应；原始差额及期间保留"),'
                   f'IF({effective}>0,"程序已确认对应",'
                   f'IF({choice}<>"","选择无效","{default}")))))')
        if not option_rows.get(item_id):
            all_records = f'COUNTIF({rng(detail,"事项编号")},{key})'
            effective_records = (f'COUNTIFS({rng(detail,"事项编号")},{key},'
                                 f'{rng(detail,"当前对应")},"?*")')
            formula = (f'=IF({effective_records}>0,'
                       f'IF({effective_records}={all_records},"原记录已由其他事项建立对应",'
                       f'"部分原记录已由其他事项建立对应，其余仍未对应"),{formula[1:]})')
        put(ix, "当前核对结论", row, formula)

    for row in range(2, 行数(detail) + 1):
        key = cell(detail, "原始记录号", row)
        lookup = f'MATCH({key},{rng(mem,"生效记录号")},0)'
        put(detail, "当前对应", row,
            f'=IFERROR(INDEX({rng(mem,"事项编号")},{lookup}),"")')
        put(detail, "当前核对编号", row,
            "=" + find(ix, "核对编号", cell(detail, "当前对应", row)))
        initial_id = str(_读值(sheets[detail], row, headers[detail]["事项编号"]))
        initial = find(ix, "当前核对结论", cell(detail, "事项编号", row))
        if option_rows.get(initial_id):
            choice = find(ix, "人工核对结果", cell(detail, "事项编号", row))
            initial = (f'IF({choice}<>"","该记录当前未对应；请按当前对应列查看",'
                       f'IF({find(ix,"核对方式",cell(detail,"事项编号",row))}="自动确认",'
                       f'"原对应已被替代，该记录当前未对应",{initial}))')
        put(detail, "当前核对结果", row,
            f'=IFERROR(INDEX({rng(mem,"显示结果")},{lookup}),{initial})')
        put(detail, "已对平金额", row,
            f'=IFERROR(INDEX({rng(mem,"已对平")},{lookup})*ABS({cell(detail,"金额",row)}),0)')
        for category in ("收入", "支出"):
            put(detail, "已确认" + category + "差", row,
                f'=IF({cell(detail,"当前对应",row)}<>"",'
                f'{cell(detail,category+"差贡献",row)},0)')

    monthly = sheets["月度核对"]
    for row in range(2, 行数("月度核对") + 1):
        month = cell("月度核对", "月份", row)
        for category in ("收入", "支出"):
            put("月度核对", "已确认" + category + "差", row,
                f'=SUMIF({rng(detail,"月份")},{month},{rng(detail,"已确认"+category+"差")})')
            put("月度核对", "未确认" + category + "差", row,
                f'={cell("月度核对",category+"金额差额",row)}'
                f'-{cell("月度核对","已确认"+category+"差",row)}')
        bal_headers = headers["月度核对"]
        if "余额差额" not in bal_headers:
            put("月度核对", "余额说明", row, "本次未提供双方可用余额")
        else:
            b = cell("月度核对", "银行-月末余额", row)
            j = cell("月度核对", "序时账-月末余额", row)
            put("月度核对", "余额说明", row,
                f'=IF(AND(ISNUMBER({b}),ISNUMBER({j})),"按双方原始余额核对","本月余额无法完整确认")')
            if row > 2:
                previous = cell("月度核对", "余额差额", row - 1)
                current = cell("月度核对", "余额差额", row)
                inc = cell("月度核对", "收入金额差额", row)
                exp = cell("月度核对", "支出金额差额", row)
                put("月度核对", "余额衔接差", row,
                    f'=IF(AND(ISNUMBER({b}),ISNUMBER({j}),ISNUMBER({previous})),'
                    f'{current}-{previous}-{inc}+{exp},"")')

    差异组成 = 簿.取表("月度差异组成")
    if 差异组成 is not None:
        hs = _列定位(差异组成)
        for row in range(2, len(差异组成.表) + 2):
            key = f'{_字母(hs["事项编号"])}{row}'
            month = f'{_字母(hs["月份"])}{row}'
            for category in ("收入", "支出"):
                confirmed = f'{_字母(hs["已确认"+category+"差"])}{row}'
                difference = f'{_字母(hs[category+"差额"])}{row}'
                _写值(差异组成, row, hs["已确认" + category + "差"],
                      f'=SUMIFS({rng(detail,"已确认"+category+"差")},'
                      f'{rng(detail,"事项编号")},{key},{rng(detail,"月份")},{month})')
                _写值(差异组成, row, hs["未确认" + category + "差"],
                      f'={difference}-{confirmed}')
            remaining = (f'COUNTIFS({rng(detail,"事项编号")},{key},{rng(detail,"月份")},{month})'
                         f'-COUNTIFS({rng(detail,"事项编号")},{key},{rng(detail,"月份")},{month},'
                         f'{rng(detail,"当前对应")},"?*")')
            _写值(差异组成, row, hs["当前核对结论"],
                  f'=IF({remaining}>0,"本月原记录仍有未对应；详见核对明细",'
                  f'"本月原记录均已建立对应，原始差额保留")')

    _结论汇总(簿, sheets, headers, rng)
    _刷主表样式(簿)


def _结论汇总(簿: 簿计划, sheets: dict, headers: dict, rng):
    """翻译 复核报表._summary_formulas：改写、追加、排序、隐藏核对结论行。"""
    tp = 簿.取表("核对结论")
    hs = _列定位(tp)
    if not {"项目", "数值"} <= hs.keys():
        return
    df = tp.表
    item_col, value_col = hs["项目"], hs["数值"]
    item_idx, value_idx = item_col - tp.前置空列 - 1, value_col - tp.前置空列 - 1
    rows = {df.iat[r, item_idx]: r + 2 for r in range(len(df))}
    原有末行 = len(df) + 1
    追加行 = set()

    def set_value(label, value):
        row = rows.get(label)
        if row is None:
            new = [None] * len(df.columns)
            new[item_idx] = label
            new[value_idx] = value
            df.loc[len(df)] = new
            row = len(df) + 1
            rows[label] = row
            追加行.add(row)
        else:
            df.iat[row - 2, value_idx] = value

    ix = "复核事项索引"
    for risk in ("低风险", "中风险", "高风险", "范围未知"):
        set_value(risk + "事项数", f'=COUNTIF({rng(ix,"风险等级")},"{risk}")')
        set_value(risk + "事项金额", f'=SUMIF({rng(ix,"风险等级")},"{risk}",{rng(ix,"影响金额")})')
    set_value("中风险抽样总数", f'=COUNTIF({rng(ix,"风险等级")},"中风险")')
    set_value("中风险抽中待核查数", f'=COUNTIF({rng(ix,"核对方式")},"人工抽样")')
    set_value("当前未对应原始记录数",
              f'=COUNTA({rng("核对明细","原始记录号")})-COUNTIF({rng("核对明细","当前对应")},"?*")')
    set_value("使用顺序", "先看月度发生额和余额，再看核对明细；人工全查及人工抽样只在黄色格选择结果。多笔业务的组成直接展开。")
    set_value("金额差额口径", "收入与支出分别比较，差额不相互抵销。人工表金额为各侧净额，差额为收入差和支出差绝对值之和；原始各笔金额见组成。")
    set_value("当前覆盖率口径", "只统计有效对应且差额为零的原记录绝对金额；已人工查看但仍未对应或仍有差额，不计作已对平。")
    set_value("月度说明", "月度始终显示双方原始发生额和余额；已确认差表示对应关系已确认，金额与记账月份均未改写；月度差异组成可查看仍未确认部分。")
    set_value("专题说明", "退款冲销重付、手续费及净额、截止性差异、重复线索及余额检查等，有事项时直接显示；月度笔数与解释分项列、技术留存表可在Excel取消隐藏查看。")
    pending = []
    for mode in ("人工全查", "人工抽样"):
        mode_range = rng(ix, "核对方式")
        choice = rng(ix, "人工核对结果")
        set_value(mode + "事项数", f'=COUNTIF({mode_range},"{mode}")')
        expression = (f'COUNTIFS({mode_range},"{mode}",{choice},"")'
                      f'+COUNTIFS({mode_range},"{mode}",{rng(ix,"当前核对结论")},"选择无效")'
                      f'+COUNTIFS({mode_range},"{mode}",{rng(ix,"当前核对结论")},"所选对应重复使用原记录")')
        pending.append(expression)
        set_value(mode + "未核对数", "=" + expression)
    set_value("待处理事项数", "=" + "+".join(pending))
    set_value("已处理事项数", f'=COUNTIF({rng(ix,"人工核对结果")},"<>")')
    set_value("人工选择冲突数", f'=COUNTIF({rng(ix,"当前核对结论")},"所选对应重复使用原记录")')
    set_value("人工结果说明", "在人工全查或人工抽样中选择结果；确认不改变原始金额和期间，否定取消对应，候选冲突不计入有效覆盖。")
    for source, label in (("银行流水", "银行"), ("序时账", "序时账")):
        numerator = f'SUMIF({rng("核对明细","来源")},"{source}",{rng("核对明细","已对平金额")})'
        明细 = sheets["核对明细"]
        金额idx = headers["核对明细"]["金额"] - 明细.前置空列 - 1
        来源idx = headers["核对明细"]["来源"] - 明细.前置空列 - 1
        total = sum(abs(float(明细.表.iat[r, 金额idx] or 0))
                    for r in range(len(明细.表))
                    if 明细.表.iat[r, 来源idx] == source)
        set_value(label + "当前已对平金额覆盖率",
                  f"={numerator}/{total}" if total else 0)
        tp.覆盖[(rows[label + "当前已对平金额覆盖率"], value_col)] = 格样式(数字格式="0.00%")
    visible = ["使用顺序", "核对范围", "范围说明", "银行有效交易笔数", "日记账有效交易笔数",
               "实际执行重要性水平", "明显微小错报临界值", "人工全查事项数", "人工全查未核对数",
               "人工抽样事项数", "人工抽样未核对数", "待处理事项数", "已处理事项数",
               "人工选择冲突数", "当前未对应原始记录数", "银行当前已对平金额覆盖率",
               "序时账当前已对平金额覆盖率", "余额核对", "期初余额差额", "期初余额状态",
               "人工结果说明", "金额差额口径", "当前覆盖率口径", "月度说明", "专题说明"]
    visible[4:4] = ["非工作日银行收付笔数", "非工作日银行收款笔数", "非工作日银行收款金额",
                    "非工作日银行付款笔数", "非工作日银行付款金额", "银行交易日历未覆盖笔数",
                    "非工作日日历口径", "非工作日日历来源"]
    set_value("已处理事项数",
              f'=COUNTIF({rng(ix,"人工核对结果")},"?*")'
              f'-COUNTIF({rng(ix,"当前核对结论")},"选择无效")'
              f'-COUNTIF({rng(ix,"当前核对结论")},"所选对应重复使用原记录")')
    contents = []
    for r in range(2, len(df) + 2):
        cov = tp.覆盖.get((r, value_col))
        # 追加行数值格在旧路径里是后建的空格，数字格式为 General 而非底层主题的 '@'；
        # 旧代码排序时把数值格的内容连同数字格式一起搬运，所以该差异随值走排序
        nf = cov.数字格式 if cov and cov.数字格式 else None
        if nf is None and r in 追加行:
            nf = "General"
        contents.append((df.iat[r - 2, item_idx], df.iat[r - 2, value_idx], nf))
        tp.覆盖.pop((r, value_col), None)
    priority = {label: number for number, label in enumerate(visible)}
    contents.sort(key=lambda item: priority.get(item[0], len(priority)))
    for row, (label, value, number_format) in enumerate(contents, 2):
        df.iat[row - 2, item_idx] = label
        df.iat[row - 2, value_idx] = value
        if number_format is not None:
            tp.覆盖[(row, value_col)] = 格样式(数字格式=number_format)
        if row > 原有末行:
            # 旧代码排序只改写项目/数值两列已存在的格：A 列空格与项目列的
            # 数字格式保持位置不动——追加区段的 A 列旧路径从未落格、项目列是
            # 后建空格（General 而非底层 '@'）
            tp.空缺格.add((row, 1))
            tp.覆盖[(row, item_col)] = 格样式(数字格式="General")
        tp.行计划取(row).隐藏 = label not in visible


_复核主表 = ("核对结论", "月度核对", "核对明细", "人工全查", "人工抽样", "每日统计")
_复核副表 = ("月度差异组成", "其他对应供选择")
_重复表 = {"疑点事项", "自动归集事项", "银行侧待查", "日记账侧待查", "逐笔匹配",
           "整组勾稽", "匹配组成", "月度统计", "其他可能对应明细"}
_技术留存 = {"运行参数", "大模型辅助明细", "解析异常明细", "运行资料与映射",
             "数据入口处置", "复核事项索引", "复核候选选项", "复核候选组成"}
_主表列宽 = {"核对编号": 10, "当前核对编号": 12, "事项编号": 24, "行别": 12, "核对原因": 28,
             "银行记录": 36, "序时账记录": 36, "银行金额": 16, "序时账金额": 16, "差额": 15,
             "人工核对结果": 24, "当前核对结论": 36, "备注": 25, "原始出处": 24, "来源": 12,
             "日期": 13, "摘要": 30, "对方及凭证": 24, "金额": 16, "程序结论": 18, "核对方式": 14,
             "当前对应": 24, "当前核对结果": 34, "月份": 12, "余额说明": 30, "项目": 29,
             "数值": 70, "摘要及对方": 38, "关系说明": 40}


def _刷主表样式(簿: 簿计划):
    """翻译 复核报表._format_sheets：表序、可见性、主表样式、打印。"""
    order = [t for n in _复核主表 if (t := 簿.取表(n)) is not None]
    簿.表 = order + [t for t in 簿.表 if all(t is not o for o in order)]
    for tp in 簿.表:
        df = tp.表
        has_rows = len(df) > 0 and any(
            v is not None and not (isinstance(v, float) and _math.isnan(v))
            for row in df.itertuples(index=False, name=None) for v in row)
        tp.状态 = ("visible" if (tp.名称 in _复核主表
                   or (tp.名称 not in _重复表 | _技术留存 and has_rows))
                   else "hidden")
        if tp.名称 not in _复核主表 and tp.名称 not in _复核副表:
            continue
        hs = _列定位(tp)
        if not hs:
            continue
        first = min(hs.values())
        while len(tp.前部列) < first - 1:
            tp.前部列.append(列计划())
        for column in range(1, first):
            tp.前部列[column - 1].隐藏 = True
        hidden = ({"原始记录号", "收入差贡献", "支出差贡献", "已对平金额",
                   "已确认收入差", "已确认支出差"} if tp.名称 == "核对明细" else set())
        if "核对编号" in hs:
            hidden.add("事项编号")
        if tp.名称 == "核对明细":
            hidden.update(("月份", "当前对应"))
        if tp.名称 in ("人工全查", "人工抽样"):
            hidden.add("填报编号")
        if tp.名称 == "月度核对":
            hidden.update(label for label in hs
                          if "笔数" in label or "变动净额" in label
                          or label.startswith("已确认") or label.startswith("未确认"))
        for label, column in hs.items():
            cp = tp.列计划取(column - tp.前置空列 - 1)
            cp.隐藏 = label in hidden
            cp.宽 = _主表列宽.get(label, 17)
        tp.网格线 = False
        tp.缩放 = 90
        tp.冻结 = f"{_字母(first + 1)}2"
        ncols = len(df.columns)
        nrows = 1 + len(df)
        tp.筛选 = f"{_字母(first)}1:{_字母(tp.前置空列 + ncols)}{max(1, nrows)}"
        h = tp.表头样式 or 格样式()
        h.字体名, h.字号, h.加粗, h.字色 = "微软雅黑", 11, True, "FFFFFF"
        h.填充色, h.填充类型 = "24445B", "solid"
        h.水平, h.垂直, h.自动换行 = "", "center", True
        tp.表头样式 = h
        # 旧 _format_sheets 表头循环覆盖 A1（无值也赋样式）但不碰边框
        tp.表头含前列 = True
        for k in range(tp.前置空列):
            tp.盖格(1, k + 1, 格样式(上框=("", None), 下框=("", None),
                                     左框=("", None), 右框=("", None)))
        # 旧 _format_sheets 对 hs 数据格整体替换字体/填充/对齐（金额/日期列含数字格式），
        # 抹掉更早层留下的同域格覆盖
        金额日期列 = {column for label, column in hs.items()
                    if (any(k in label for k in ("金额", "收入", "支出", "余额", "差额", "发生额差"))
                        and label != "余额说明") or label == "日期"}
        hs列集合 = set(hs.values())
        for (r, c), cov in list(tp.覆盖.items()):
            if r < 2 or c not in hs列集合:
                continue
            cov.字体名 = cov.字号 = cov.加粗 = cov.字色 = cov.下划线 = None
            cov.填充色 = cov.填充类型 = None
            cov.水平 = cov.垂直 = cov.自动换行 = None
            if c in 金额日期列:
                cov.数字格式 = None
            if cov.属性元组() == (None,) * len(cov.属性元组()):
                del tp.覆盖[(r, c)]
        tp.行计划取(1).高 = 36
        for ri in range(len(df)):
            row = ri + 2
            height = 32
            rp = tp.行计划取(row)
            rp.样式 = 合并样式(rp.样式, 格样式(
                字体名="微软雅黑", 字号=10, 加粗=False, 字色="243746",
                填充色="F1F5F8" if row % 2 else "FFFFFF", 填充类型="solid",
                水平="", 垂直="top", 自动换行=True))
            for label, column in hs.items():
                j = column - tp.前置空列 - 1
                base = tp.列计划取(j).基础样式 or 格样式()
                if any(k in label for k in ("金额", "收入", "支出", "余额", "差额", "发生额差")) \
                        and label != "余额说明":
                    base.数字格式 = "#,##0.00;[Red]-#,##0.00"
                if label == "日期":
                    base.数字格式 = "yyyy-mm-dd"
                tp.列计划取(j).基础样式 = base
                v = df.iat[ri, j]
                if not (isinstance(v, str) and v.startswith("=")) and label not in hidden:
                    width = _主表列宽.get(label, 17)
                    lines = sum(
                        max(1, (_中文宽(part) + width - 1) // width)
                        for part in str(v or "").split("\n"))
                    height = max(height, lines * 15 + 5)
                if label == "人工核对结果" and tp.名称 in ("人工全查", "人工抽样") \
                        and df.iat[ri, hs["行别"] - tp.前置空列 - 1] == "核对事项":
                    tp.盖格(row, column, 格样式(填充色="FFF2CC", 填充类型="solid"))
            rp.高 = min(height, 300)
        tp.打印 = {"orientation": "landscape", "paperSize": 8, "fitToWidth": 1,
                   "fitToHeight": 0, "fitToPage": True, "title_rows": "1:1",
                   "area": f"{_字母(first)}1:{_字母(tp.前置空列 + ncols)}{nrows}",
                   "左边距": 0.25, "右边距": 0.25}
    簿.active = 0
    簿.自动重算 = True


# ═══════════════════════════════════════════════════════════════════════════
# 展示层回放（翻译自 报告展示.format_display_sheets / _style / _monthly）
# ═══════════════════════════════════════════════════════════════════════════

_展示金额列 = ("银行金额", "序时账金额", "金额", "差额")


def 应用展示层(簿: 簿计划):
    """翻译 报告展示.format_display_sheets。"""
    from 排版共享 import 阅读表, 组成金额列

    references = {}
    for name in ("核对明细", "复核事项索引", "人工全查", "人工抽样"):
        t = 簿.取表(name)
        if t is None:
            raise KeyError(name)
        references[name] = (_列定位(t), max(2, len(t.表) + 1))

    def find(name, key_label, key, label):
        headers, end = references[name]
        value_col = _字母(headers[label])
        key_col = _字母(headers[key_label])
        return (f"IFERROR(INDEX('{name}'!${value_col}$2:${value_col}${end},"
                f"MATCH({key},'{name}'!${key_col}$2:${key_col}${end},0)),\"\")")

    manual_links = {}
    for name in ("人工全查", "人工抽样"):
        tp = 簿.取表(name)
        hs = _列定位(tp)
        for row in range(2, len(tp.表) + 2):
            if _读值(tp, row, hs["填报编号"]):
                manual_links[str(_读值(tp, row, hs["事项编号"]))] = (
                    name, row, hs["人工核对结果"])

    for name in ("逐笔核对", "整组核对", "未对应记录", "人工全查", "人工抽样"):
        tp = 簿.取表(name)
        hs = _列定位(tp)
        for row in range(2, len(tp.表) + 2):
            item_coord = f"{_字母(hs['事项编号'])}{row}"
            current = find("复核事项索引", "事项编号", item_coord, "当前核对结论")
            if name == "逐笔核对":
                keys = [f"{_字母(hs[label])}{row}" for label in ("银行记录号", "序时账记录号")]
                owners = [find("核对明细", "原始记录号", key, "当前对应") for key in keys]
                numbers = [find("复核事项索引", "事项编号", owner, "核对编号") for owner in owners]
                bank_location = f'IF({owners[0]}="","未对应","已确认对应（核对"&{numbers[0]}&"）")'
                journal_location = f'IF({owners[1]}="","未对应","已确认对应（核对"&{numbers[1]}&"）")'
                current = (f'IF(AND({owners[0]}={owners[1]},{owners[0]}<>""),'
                           f'IF({owners[0]}={item_coord},{current},'
                           f'"已并入核对"&{numbers[0]}&"，对应已确认"),'
                           f'IF(AND({owners[0]}="",{owners[1]}=""),{current},'
                           f'"银行："&{bank_location}&"；序时账："&{journal_location}))')
            if name == "整组核对" or (
                    name in ("人工全查", "人工抽样")
                    and _读值(tp, row, hs["行别"]) == "核对事项"
                    and all(_读值(tp, row, hs[label]) is None for label in 组成金额列)):
                current = f'SUBSTITUTE({current},"原对应被人工选择替代，记录未对应","原对应已被替代；展开查看各笔当前去向")'
                current = f'SUBSTITUTE({current},"原对应已否定，记录恢复未对应","原对应已否定；展开查看各笔当前去向")'
            elif name in ("人工全查", "人工抽样") and _读值(tp, row, hs["行别"]) == "核对事项":
                current = f'SUBSTITUTE({current},"原对应被人工选择替代，记录未对应","原对应已被替代；当前去向见逐笔核对")'
                current = f'SUBSTITUTE({current},"原对应已否定，记录恢复未对应","原对应已否定；当前去向见逐笔核对")'
            if "当前核对结论" in hs and (name == "逐笔核对"
                    or _读值(tp, row, hs["行别"]) == "核对事项"):
                _写值(tp, row, hs["当前核对结论"], "=" + current)
            if "原始记录号" in hs and _读值(tp, row, hs["原始记录号"]):
                target = "当前核对结果" if name == "未对应记录" else "当前记录结果"
                key = f"{_字母(hs['原始记录号'])}{row}"
                record_result = "=" + find("核对明细", "原始记录号", key, "当前核对结果")
                _写值(tp, row, hs[target], record_result)
                owner = find("核对明细", "原始记录号", key, "当前对应")
                number = find("核对明细", "原始记录号", key, "当前核对编号")
                display = (f'=IF(AND({owner}<>"",{owner}<>{item_coord}),'
                           f'"已并入核对"&{number}&"，对应已确认",{record_result[1:]})')
                if "当前核对结论" in hs:
                    _写值(tp, row, hs["当前核对结论"], display)
                else:
                    _写值(tp, row, hs[target], display)
            if "行别" in hs and "其他对应" in str(_读值(tp, row, hs["行别"])):
                _写值(tp, row, hs["当前核对结论"], None)
            if name not in ("人工全查", "人工抽样") \
                    and str(_读值(tp, row, hs["事项编号"])) in manual_links:
                dest, dest_row, dest_col = manual_links[str(_读值(tp, row, hs["事项编号"]))]
                dest_headers, dest_end = references[dest]
                key_col = _字母(dest_headers["填报编号"])
                label = _读值(tp, row, hs["核对编号"])
                _写值(tp, row, hs["核对编号"],
                      f'=HYPERLINK("#\'{dest}\'!"&ADDRESS(MATCH({item_coord},'
                      f"'{dest}'!${key_col}$2:${key_col}${dest_end},0)+1,{dest_col},4),\"{label}\")")
        _展示样式(tp)
    _展示样式(簿.取表("非工作日交易"))
    _月度(簿.取表("月度核对"))

    summary = 簿.取表("核对结论")
    hs = _列定位(summary)
    messages = {
        "使用顺序": "先看非工作日银行收付示警，再看月度收支与余额、逐笔或整组。整组左侧“＋”展开原始记录。人工全查、人工抽样直接在黄色格选择结果。",
        "金额差额口径": "收入与支出分别比较。整组汇总金额写在文字格；收入、支出数值列只记录原始组成，避免重复加总。逐笔金额正数为收入、负数为支出。",
        "月度说明": "月度收支和余额在同一张表。人工确认不改变原始金额、记账月份和客观差额；年度合计只加总发生额。",
        "专题说明": "非工作日交易按银行所选日期独立示警；每日、退款重付、手续费、跨期、重复线索及输入详情保留在隐藏表中。",
        "人工结果说明": "黄色格每项只选一次并保存。点击逐笔、整组或未对应表的核对编号，可跳到人工填写处。原始分组保留，当前结果随选择更新。",
    }
    for row in range(2, len(summary.表) + 2):
        label = _读值(summary, row, hs["项目"])
        if label in messages:
            _写值(summary, row, hs["数值"], messages[label])
    summary.列计划取(hs["项目"] - summary.前置空列 - 1).宽 = 28
    summary.列计划取(hs["数值"] - summary.前置空列 - 1).宽 = 100
    for row in range(2, len(summary.表) + 2):
        value = _读值(summary, row, hs["数值"])
        is_formula = isinstance(value, str) and value.startswith("=")
        lines = max(1, (len(str(value or "")) * 2 + 99) // 100) if not is_formula else 1
        summary.行计划取(row).高 = max(25, lines * 15 + 6)

    order = []
    for n in 阅读表:
        t = 簿.取表(n)
        if t is None:
            raise KeyError(n)
        order.append(t)
    簿.表 = order + [t for t in 簿.表 if all(t is not o for o in order)]
    for tp in 簿.表:
        tp.状态 = "visible" if tp.名称 in 阅读表 else "hidden"
    簿.active = 0


_展示隐藏列 = {"事项编号", "原始记录号", "银行记录号", "序时账记录号", "填报编号", "当前记录结果"}
_展示列宽 = {"核对编号": 9, "行别": 12, "银行日期": 12, "序时账日期": 12, "日期": 12,
             "交易日期": 14, "原日期列名": 18, "原日期值": 18, "原文件行号": 14,
             "休息日类型": 16, "节日名称": 18, "收支方向": 12, "对方及辅助文字": 36,
             "银行记录": 34, "序时账记录": 34, "银行金额": 14, "序时账金额": 14, "金额": 14,
             "差额": 13, "核对差额": 22, "人工核对结果": 24, "当前核对结论": 29,
             "当前记录结果": 27, "当前核对结果": 30, "核对原因": 36, "核对依据": 36,
             "核对方式": 12, "摘要": 30, "对方及凭证": 27, "原始出处": 28,
             "银行原始出处": 26, "序时账原始出处": 26, "备注": 25}


def _展示样式(tp: 表计划):
    """翻译 报告展示._style。"""
    from 排版共享 import 组成金额列

    hs = _列定位(tp)
    df = tp.表
    偏 = tp.前置空列
    ncols = len(df.columns)
    nrows = 1 + len(df)
    manual = tp.名称 in ("人工全查", "人工抽样")
    grouped = "行别" in hs
    hidden = set(_展示隐藏列)
    if manual:
        hidden.update(("银行金额", "序时账金额", "差额"))
    while len(tp.前部列) < min(hs.values()) - 1:
        tp.前部列.append(列计划())
    for col in range(1, min(hs.values())):
        tp.前部列[col - 1].隐藏 = True
    for label, col in hs.items():
        cp = tp.列计划取(col - 偏 - 1)
        cp.隐藏 = label in hidden
        cp.宽 = _展示列宽.get(label, 14)
        base = cp.基础样式 or 格样式()
        if label in 组成金额列 or label in _展示金额列:
            base.数字格式 = "#,##0.00;[Red]-#,##0.00"
        if label.endswith("日期"):
            base.数字格式 = "yyyy-mm-dd"
        cp.基础样式 = base
    tp.网格线 = False
    tp.缩放 = 85
    tp.冻结 = "C2" if grouped else "B2"
    tp.outline汇总在下 = False
    tp.显示outline符号 = True
    tp.筛选 = None if grouped else f"A1:{_字母(偏 + ncols)}{nrows}"
    tp.行计划取(1).高 = 32
    h = tp.表头样式 or 格样式()
    h.字体名, h.字号, h.加粗, h.字色 = "微软雅黑", 10, True, "FFFFFF"
    h.填充色, h.填充类型 = "24445B", "solid"
    h.水平, h.垂直, h.自动换行 = "", "center", True
    tp.表头样式 = h
    tp.表头含前列 = True
    # 旧 _style 给 A1 上字体/填充/对齐但不碰边框；表头样式带着主题边框，
    # 前置空列表头需显式清掉边框才与旧产物一致
    for k in range(偏):
        tp.盖格(1, k + 1, 格样式(上框=("", None), 下框=("", None),
                                 左框=("", None), 右框=("", None)))

    last_main = None
    for ri in range(len(df)):
        row = ri + 2
        kind = _读值(tp, row, hs["行别"]) if grouped else ""
        main = kind == "核对事项"
        other = "其他对应" in str(kind)
        if main:
            last_main = row
        child = grouped and not main
        rp = tp.行计划取(row)
        rp.outline = int(child)
        rp.隐藏 = bool(child and tp.名称 == "整组核对")
        rp.收起 = False
        if child and last_main and tp.名称 == "整组核对":
            tp.行计划取(last_main).收起 = True
        rp.样式 = 合并样式(rp.样式, 格样式(
            字体名="微软雅黑", 字号=10, 加粗=main, 字色="243746",
            填充色="DEEAF2" if main else "FFF8E8" if other
                  else "F3F7FA" if kind == "银行组成" else "FFFFFF",
            填充类型="solid", 水平="", 垂直="center", 自动换行=True,
            **({"上框": ("thin", "91A8B8"), "下框": ("", None),
                "左框": ("", None), "右框": ("", None)} if main else {})))
        height = 40
        for label, col in hs.items():
            j = col - 偏 - 1
            v = df.iat[ri, j]
            if label == "人工核对结果" and manual and main:
                tp.盖格(row, col, 格样式(填充色="FFF2CC", 填充类型="solid"))
            elif label == "人工核对结果" and child:
                tp.盖格(row, col, 格样式(数字格式=";;;"))
            if isinstance(v, str) and v.startswith("=HYPERLINK("):
                tp.盖格(row, col, 格样式(字体名="微软雅黑", 字号=10, 加粗=False,
                                         字色="0563C1", 下划线="single"))
            if label not in hidden and not (isinstance(v, str) and v.startswith("=")):
                width = _展示列宽.get(label, 14)
                lines = sum(max(1, int((_中文宽(part) + width - 1) // width))
                            for part in str(v or "").split("\n"))
                height = max(height, lines * 15 + 8)
        rp.高 = min(height, 350)
    for label in ("银行金额", "序时账金额", "金额"):
        if label in hs and label not in hidden:
            tp.盖格(1, hs[label], 格样式(批注=("正数为收入，负数为支出。", "核对说明")))
    if grouped:
        tp.盖格(1, hs["银行记录"], 格样式(批注=(
            "汇总行先看双方笔数、收入与支出；组成行先银行后序时账。两侧排在同一组不表示已经逐笔配对。",
            "核对说明")))
    tp.打印 = {"orientation": "landscape", "paperSize": 8, "fitToWidth": 1,
               "fitToHeight": 0, "fitToPage": True, "title_rows": "1:1",
               "area": f"A1:{_字母(max(c for k, c in hs.items() if k not in hidden))}{nrows}",
               "左边距": 0.2, "右边距": 0.2}


def _月度(tp: 表计划):
    """翻译 报告展示._monthly。"""
    hs = _列定位(tp)
    df = tp.表
    偏 = tp.前置空列
    ncols = len(df.columns)
    end = len(df) + 1
    balance_cols = [hs[x] for x in ("银行-月末余额", "序时账-月末余额", "余额差额")]
    for row in range(2, end + 1):
        for col in balance_cols:
            if _读值(tp, row, col) is None:
                _写值(tp, row, col, "缺失")
                tp.盖格(row, col, 格样式(字体名="微软雅黑", 字号=10, 加粗=False, 字色="9C6500"))
    if end >= 2:
        for label in ("收入金额差额", "支出金额差额", "余额差额", "余额衔接差"):
            letter = _字母(hs[label])
            tp.条件格式.append((
                f"{letter}2:{letter}{end}",
                条件规则(公式=[f"AND(ISNUMBER({letter}2),{letter}2<>0)"],
                         填充色="FFF2CC", 字色="9C0006")))
    years = sorted({str(_读值(tp, row, hs["月份"]))[:4] for row in range(2, end + 1)})
    for year in years:
        rows = [r for r in range(2, end + 1)
                if str(_读值(tp, r, hs["月份"])).startswith(year + "-")]
        if not rows:
            continue
        new = [None] * ncols
        nf_cols = []
        for label, col in hs.items():
            j = col - 偏 - 1
            if "金额" in label or "笔数" in label or "变动净额" in label \
                    or label.startswith(("已确认", "未确认")):
                new[j] = "=" + "+".join(f"{_字母(col)}{r}" for r in rows)
                nf_cols.append(col)
        new[hs["月份"] - 偏 - 1] = year + "年合计"
        new[hs["余额说明"] - 偏 - 1] = "发生额合计；余额不加总"
        df.loc[len(df)] = new
        target = len(df) + 1
        # 合计行是后加的格子：旧路径非求和列保持默认 General，不继承底层主题格式
        for label, col in hs.items():
            if col not in nf_cols:
                tp.盖格(target, col, 格样式(数字格式="General"))
        for col in nf_cols:
            tp.盖格(target, col, 格样式(数字格式="#,##0.00;[Red]-#,##0.00"))
        rp = tp.行计划取(target)
        rp.高 = 32
        rp.样式 = 合并样式(rp.样式, 格样式(
            字体名="微软雅黑", 字号=10, 加粗=True, 字色="243746",
            填充色="DEEAF2", 填充类型="solid", 水平="", 垂直="center", 自动换行=True))
        # 旧路径合计行样式也落到了前置空列 A 格（渲染器行样式不落 A 列，显式补盖）
        if 偏:
            tp.盖格(target, 1, 格样式(
                字体名="微软雅黑", 字号=10, 加粗=True, 字色="243746",
                填充色="DEEAF2", 填充类型="solid", 垂直="center", 自动换行=True))
    tp.筛选 = f"A1:{_字母(偏 + ncols)}{end}"
    tp.打印 = dict(tp.打印)
    tp.打印["area"] = f"A1:{_字母(偏 + ncols)}{len(df) + 1}"
    tp.冻结 = "B2"
    for label, col in hs.items():
        tp.列计划取(col - 偏 - 1).宽 = 13 if label != "余额说明" else 28


# ═══════════════════════════════════════════════════════════════════════════
# 紧凑层回放（翻译自 紧凑报告.apply_compact_report / _write_items /
# _period_flags / _overview / format_compact_sheet；逐条对照源逻辑）
# ═══════════════════════════════════════════════════════════════════════════

import re as _re
import unicodedata as _uni
from copy import copy as _copy
from datetime import date as _date, datetime as _dt
from decimal import Decimal as _Decimal

_紧凑阅读表 = ("核对总览", "待处理事项", "月度核对", "每日统计", "对应明细")
_紧凑阅读色 = "BDD7EE"
_紧凑留痕色 = "C6E0B4"
_紧凑人工方式 = ("人工全查", "人工抽样")
_紧凑列宽 = {"核对编号": 7, "处理要求": 9, "对应方式": 9, "哪里需要查看": 26,
             "银行记录": 28, "序时账记录": 28, "人工核对结果": 19, "当前结果": 21,
             "备注": 20, "完整明细": 7, "核对依据": 26, "日期": 11, "月份": 10,
             "需查看": 13, "余额说明": 21, "项目": 27, "结果": 43, "怎么看": 47}


def _查找(簿: 簿计划, name, label, key, key_label="事项编号", default='""'):
    """紧凑报告._lookup 的计划层等价。"""
    t = 簿.取表(name)
    hs = _列定位(t)
    end = max(2, len(t.表) + 1)
    vcol, kcol = (_字母(hs[x]) for x in (label, key_label))
    return (f"IFERROR(INDEX('{name}'!${vcol}$2:${vcol}${end},"
            f"MATCH({key},'{name}'!${kcol}$2:${kcol}${end},0)),{default})")


def _text计划(tp: 表计划, row, col, text, limit=48):
    """紧凑报告._text 的计划层等价：截断+批注+公式符号前置引号。"""
    text = str(text or "").replace("\r", "").replace("\n", "；")
    value = text if len(text) <= limit else text[:limit - 1] + "…"
    if len(text) > limit:
        tp.盖格(row, col, 格样式(批注=(text, "完整说明")))
    if value and value[0] in "=+-@":
        value = "'" + value
    _写值(tp, row, col, value)


def _行号(coord):
    return int(_re.sub(r"[^0-9]", "", coord))


def _sqref行(sqref):
    rows = []
    for part in str(sqref).split():
        if ":" in part:
            a, b = part.split(":", 1)
            rows.extend(range(_行号(a), _行号(b) + 1))
        elif part:
            rows.append(_行号(part))
    return rows


def _含坐标(sqref, coord):
    col = _re.sub(r"[0-9]", "", coord)
    row = _行号(coord)
    for part in str(sqref).split():
        if ":" in part:
            a, b = part.split(":", 1)
            ca, cb = _re.sub(r"[0-9]", "", a), _re.sub(r"[0-9]", "", b)
            if ca <= col <= cb and _行号(a) <= row <= _行号(b):
                return True
        elif part == coord:
            return True
    return False


def _写事项(簿: 簿计划):
    """翻译 紧凑报告._write_items：建待处理事项/对应明细，回收人工表入口。"""
    from 排版共享 import (记录金额 as _amounts, 记录描述 as _description,
                          入选原因 as _plain_reason, 短结果 as _short_result)

    明细 = 簿.取表("核对明细")
    hs_明细 = _列定位(明细)
    records = _defaultdict(lambda: {"银行流水": [], "序时账": []})
    first_record = {}
    for row in range(2, len(明细.表) + 2):
        item = {label: _读值(明细, row, col) for label, col in hs_明细.items()}
        records[item["事项编号"]][item["来源"]].append(item)
        first_record.setdefault(item["事项编号"], row)
    manual = {}
    for name in _紧凑人工方式:
        tp = 簿.取表(name)
        hs = _列定位(tp)
        for row in range(2, len(tp.表) + 2):
            r = {label: _读值(tp, row, col) for label, col in hs.items()}
            if r.get("行别") == "核对事项":
                manual[r["事项编号"]] = (name, row, hs, r)

    ix = 簿.取表("复核事项索引")
    hs_ix = _列定位(ix)
    items = [{label: _读值(ix, row, col) for label, col in hs_ix.items()}
             for row in range(2, len(ix.表) + 1 + 1)]
    priority = {"人工全查": 0, "人工抽样": 1, "留存备查": 2, "自动确认": 3}
    items.sort(key=lambda it: (priority.get(it["核对方式"], 2), int(it["核对编号"])))
    numbers = {it["事项编号"]: it["核对编号"] for it in items}
    end = max(2, len(items) + 1)

    待处理列 = ["核对编号", "处理要求", "哪里需要查看", "银行记录", "序时账记录",
                "收入差额", "支出差额", "人工核对结果", "当前结果", "备注",
                "完整明细", "事项编号"]
    对应列 = ["核对编号", "对应方式", "银行记录", "银行收入", "银行支出",
              "序时账记录", "序时账收入", "序时账支出", "收入差额", "支出差额",
              "当前结果", "核对依据", "完整明细", "事项编号"]
    pending_data, pending_extra = [], []
    matched_data, matched_extra = [], []
    pending_rows = {}
    for it in items:
        identity = it["事项编号"]
        bank = records[identity]["银行流水"]
        journal = records[identity]["序时账"]
        bi, be = _amounts(bank)
        ji, je = _amounts(journal)
        mode = it["核对方式"]
        if mode != "自动确认":
            pending_data.append([
                it["核对编号"],
                {"人工全查": "逐项核对", "人工抽样": "抽中核对"}.get(mode, "留存查看"),
                _plain_reason(it, bank, journal, bi - ji, be - je),
                _description(bank), _description(journal), bi - ji, be - je,
                "", "", "", "查看", identity])
            row = len(pending_data) + 1
            pending_rows[identity] = row
            pending_extra.append((it, row))
        if bank and journal:
            matched_data.append([
                it["核对编号"],
                "逐笔对应" if len(bank) == len(journal) == 1 else "整组对应",
                _description(bank), bi, be, _description(journal), ji, je,
                bi - ji, be - je, "", it.get("判断依据") or "", "查看", identity])
            matched_extra.append((it, len(bank) == len(journal) == 1,
                                  len(matched_data) + 1))

    待处理 = 表计划(名称="待处理事项", 表=pd.DataFrame(pending_data, columns=待处理列))
    待处理.物化空格 = True      # 旧路径紧凑层在重载后直接建表，空串格也带紧凑样式落格
    簿.表.append(待处理)
    对应 = 表计划(名称="对应明细", 表=pd.DataFrame(matched_data, columns=对应列))
    对应.物化空格 = True
    簿.表.append(对应)

    for it, row in pending_extra:
        identity = it["事项编号"]
        for col in (3, 4, 5):
            _text计划(待处理, row, col, 待处理.表.iat[row - 2, col - 1], 32)
        待处理.盖格(row, 3, 格样式(批注=(
            "；".join(str(it.get(k) or "") for k in ("入选原因", "判断依据")),
            "原判断依据")))
        key = f"$L{row}"
        _写值(待处理, row, 9,
              _short_result(_查找(簿, "复核事项索引", "当前核对结论", key)))
        target = manual.get(identity)
        if target:
            name, source_row, hs_m, original = target
            _写值(待处理, row, 8, original.get("人工核对结果") or None)
            _写值(待处理, row, 10, original.get("备注") or None)
            人工tp = 簿.取表(name)
            old_coord = f"{_字母(hs_m['人工核对结果'])}{source_row}"
            for v in 人工tp.验证:
                if _含坐标(v.get("sqref", ""), old_coord):
                    nv = dict(v)
                    nv["sqref"] = f"H{row}"
                    待处理.验证.append(nv)
            owner_key = f"{_字母(hs_m['填报编号'])}{source_row}"
            for label, col_l in (("人工核对结果", "H"), ("备注", "J")):
                value = (f"IFERROR(INDEX('待处理事项'!${col_l}$2:${col_l}${end},"
                         f"MATCH({owner_key},'待处理事项'!$L$2:$L${end},0)),\"\")")
                _写值(人工tp, source_row, hs_m[label], f'=IF({value}="","",{value})')
        else:
            _写值(待处理, row, 8, "本次无需填写")
        target_row = first_record.get(identity, 2)
        link = (f"#'{target[0]}'!A{target[1]}" if target
                else f"#'核对明细'!A{target_row}")
        待处理.盖格(row, 11, 格样式(超链接=link))

    for it, single, row in matched_extra:
        identity = it["事项编号"]
        for col in (3, 6, 12):
            _text计划(对应, row, col, 对应.表.iat[row - 2, col - 1], 32)
        key = f"$N{row}"
        _写值(对应, row, 11,
              _short_result(_查找(簿, "复核事项索引", "当前核对结论", key)))
        if identity in pending_rows:
            _写值(对应, row, 1,
                  f'=HYPERLINK("#\'待处理事项\'!H"&(MATCH({key},'
                  f"'待处理事项'!$L$2:$L${end},0)+1),\"{it['核对编号']}\")")
        detail_name = "逐笔核对" if single else "整组核对"
        dt = 簿.取表(detail_name)
        hs_d = _列定位(dt)
        dkey = _字母(hs_d["事项编号"])
        dend = max(2, len(dt.表) + 1)
        _写值(对应, row, 13,
              f'=HYPERLINK("#\'{detail_name}\'!A"&(MATCH({key},'
              f"'{detail_name}'!${dkey}$2:${dkey}${dend},0)+1),\"查看\")")

    for name in _紧凑人工方式:
        簿.取表(name).验证 = []
    for name in ("逐笔核对", "整组核对", "未对应记录"):
        tp = 簿.取表(name)
        hs = _列定位(tp)
        for row in range(2, len(tp.表) + 2):
            item_id = _读值(tp, row, hs["事项编号"])
            if item_id not in pending_rows:
                continue
            key = f"{_字母(hs['事项编号'])}{row}"
            number = str(_读值(tp, row, hs["核对编号"]))
            if number.startswith("="):
                number = numbers[item_id]
            _写值(tp, row, hs["核对编号"],
                  f'=HYPERLINK("#\'待处理事项\'!H"&(MATCH({key},'
                  f"'待处理事项'!$L$2:$L${end},0)+1),\"{number}\")")


def _期间标记(簿: 簿计划, name):
    """翻译 紧凑报告._period_flags。"""
    tp = 簿.取表(name)
    daily = name == "每日统计"
    balance_labels = (("银行-期末余额", "日记账-期末余额") if daily
                      else ("银行-月末余额", "序时账-月末余额"))
    追加列 = []
    for label in (*balance_labels, "余额差额"):
        if label not in _列定位(tp):
            tp.表[label] = "未提供"
            追加列.append(label)
    tp.表["需查看"] = None
    追加列.append("需查看")
    hs = _列定位(tp)
    anomaly_days = set()
    异常 = 簿.取表("余额连续性异常")
    if 异常 is not None:
        hs_a = _列定位(异常)
        for row in range(2, len(异常.表) + 2):
            for label in ("日期", "异常日期", "date"):
                if label in hs_a:
                    v = _读值(异常, row, hs_a[label])
                    if v:
                        anomaly_days.add(str(v)[:10])
    diff_labels = [l for l in ("收入金额差额", "支出金额差额", "余额差额", "余额衔接差")
                   if l in hs]
    column = hs["需查看"]
    for row in range(2, len(tp.表) + 2):
        period = _读值(tp, row, hs["日期" if daily else "月份"])
        if not daily and "合计" in str(period):
            continue
        tests = [f"AND(ISNUMBER({_字母(hs[l])}{row}),{_字母(hs[l])}{row}<>0)"
                 for l in diff_labels]
        missing = [f"NOT(ISNUMBER({_字母(hs[l])}{row}))" for l in balance_labels]
        if (daily and str(period)[:10] in anomaly_days) or (
                not daily
                and str(period)[:7] in {day[:7] for day in anomaly_days}):
            _写值(tp, row, column, "余额不连续")
        else:
            _写值(tp, row, column,
                  f'=IF(OR({",".join(tests)}),"有差额",'
                  f'IF(OR({",".join(missing)}),"余额未提供",""))')
    # 追加列在旧路径里是紧凑层才出现的新列：不背复核层行样式，中和为默认格
    # （有值的格随后由紧凑变换再盖一遍）；合计行被跳过的需查看格旧路径不存在。
    # 中和必须连边框一起清掉（复核层行样式的 medium/dashed 边框旧路径不存在），
    # 表头格同理：旧路径新列表头是空格后盖紧凑表头，不带底层边框。
    清边框 = dict(上框=("", None), 下框=("", None), 左框=("", None), 右框=("", None))
    for label in 追加列:
        col = hs[label]
        tp.盖格(1, col, 格样式(**清边框))
        for row in range(2, len(tp.表) + 2):
            if _读值(tp, row, col) is None:
                tp.空缺格.add((row, col))
                continue
            tp.盖格(row, col, 格样式(字体名="Calibri", 字号=11, 字色="theme:1",
                                     填充类型="none", 垂直="", 自动换行=False,
                                     数字格式="General", **清边框))


def _建总览(簿: 簿计划):
    """翻译 紧凑报告._overview。"""
    结论 = 簿.取表("核对结论")
    hs_c = _列定位(结论)
    rows = [
        ("核对范围", "核对范围", "资料有缺口时，结果要结合原始说明查看。"),
        ("资料情况", "范围说明", ""),
        ("银行记录数", "银行有效交易笔数", "进入本次核对的银行交易笔数。"),
        ("账上记录数", "日记账有效交易笔数", "进入本次核对的银行存款记账笔数。"),
        ("银行金额已对平比例", "银行当前已对平金额覆盖率", "收付金额分开累计；只计入已对应且差额为零的记录。"),
        ("账上金额已对平比例", "序时账当前已对平金额覆盖率", "人工查看过但仍有差额，不算已对平。"),
        ("还需人工核对", "待处理事项数", "在待处理事项的黄色格选择结果；留存查看不计入待办。"),
        ("已完成人工核对", "已处理事项数", "已选有效结果的事项数，有差额的仍保留差额。"),
        ("人工选择有冲突", "人工选择冲突数", "同一原记录被重复采用时，需要重新选择。"),
        ("还没找到对应的记录", "当前未对应原始记录数", "银行和账上合计笔数；未对应不等于漏记。"),
        ("期初余额差额", "期初余额差额", "银行余额减账上余额；未提供可用余额时不补零。"),
        ("余额核对情况", "余额核对", "每日和月度表保留收支、余额及差额。"),
    ]
    if 簿.取表("非工作日交易") is not None:
        rows.extend([
            ("非工作日银行收付笔数", "非工作日银行收付笔数", "已对平的记录也保留提示；完整清单在绿色非工作日交易表。"),
            ("非工作日银行收款笔数", "非工作日银行收款笔数", ""),
            ("非工作日银行收款金额", "非工作日银行收款金额", ""),
            ("非工作日银行付款笔数", "非工作日银行付款笔数", ""),
            ("非工作日银行付款金额", "非工作日银行付款金额", ""),
            ("银行交易日历未覆盖笔数", "银行交易日历未覆盖笔数", "超出2021至2026年日历的记录，不能据此判断是否为工作日。"),
        ])
    out = []
    for title, label, explanation in rows:
        value = _查找(簿, "核对结论", "数值", '"' + label + '"', "项目", '"未提供"')
        if label == "核对范围":
            for old, new in (("范围受限", "资料有缺口，部分结果需核实"),
                             ("范围可用", "本次资料可以用于核对"),
                             ("范围未验证", "尚未确认资料范围")):
                value = f'SUBSTITUTE({value},"{old}","{new}")'
        if label == "范围说明":
            original = next(
                (_读值(结论, r, hs_c["数值"])
                 for r in range(2, len(结论.表) + 2)
                 if _读值(结论, r, hs_c["项目"]) == label), "")
            value = (f'IF({value}="未发现范围疑点","未发现额外资料问题",'
                     f'"有需要确认的资料，见原始说明")')
            explanation = str(original or "")
        if label == "余额核对":
            value = f'SUBSTITUTE({value},"余额核对未实施","未提供双方可用余额")'
        out.append([title, "=" + value, explanation])
    tp = 表计划(名称="核对总览", 表=pd.DataFrame(out, columns=["项目", "结果", "怎么看"]))
    tp.物化空格 = True          # 旧路径紧凑层在重载后直接建表，空串格也带紧凑样式落格
    簿.表.append(tp)
    for i, (title, label, explanation) in enumerate(rows):
        row = i + 2
        _text计划(tp, row, 3, tp.表.iat[row - 2, 2], 40)
        if "比例" in title:
            tp.盖格(row, 2, 格样式(数字格式="0.00%"))
        elif "差额" in title or "金额" in title:
            tp.盖格(row, 2, 格样式(数字格式="#,##0.00;[Red]-#,##0.00"))
    for name, title in (("月度核对", "有差额的月份"), ("每日统计", "有差额或余额不连续的日期")):
        t = 簿.取表(name)
        column = _字母(_列定位(t)["需查看"])
        rng = f"'{name}'!${column}$2:${column}${max(2, len(t.表) + 1)}"
        row = len(tp.表) + 2
        tp.表.loc[len(tp.表)] = [title,
                                 f'=COUNTIF({rng},"有差额")+COUNTIF({rng},"余额不连续")',
                                 "打开" + name + "，筛选“需查看”列。"]
        tp.盖格(row, 3, 格样式(超链接=f"#'{name}'!A1"))
    search = _查找(簿, "核对结论", "数值", '"组合搜索完整性"', "项目")
    tp.表.loc[len(tp.表)] = [
        "对应查找情况",
        f'=IF({search}="范围受限","部分查找未完成",'
        f'IF({search}="存在多解","有多种可能对应","已在程序设定的范围内查找"))',
        "未找到不等于没有对应，详细检查范围见绿色运行参数表。"]
    tp.表.loc[len(tp.表)] = [
        "使用方法", "先看总览和每日、月度差额，再处理黄色格。",
        "浅蓝色表用于阅读；浅绿色表保留完整记录和计算依据。"]


def _做变换(reading):
    """format_compact_sheet 数据格部分：在渲染器每格合并后调用。"""
    def 变换(样式, 值, 列标签):
        # 旧代码 copy(cell.font)：从未显式赋过字体的格子带默认主题色，
        # 计划层以 字色 None 表示，这里还原为 theme:1 供渲染器落成 Color(theme=1)
        无字体痕迹 = all(getattr(样式, k) is None
                       for k in ("字体名", "字号", "加粗", "字色", "下划线"))
        样式.字体名, 样式.字号 = "MiSans", 9
        if 无字体痕迹:
            样式.字色 = "theme:1"
        样式.填充类型, 样式.填充色 = "none", None
        numeric = isinstance(值, (int, float, _Decimal)) and not isinstance(值, bool)
        样式.水平 = "right" if numeric else "left"
        样式.垂直 = "top"
        样式.自动换行 = reading
        if numeric and any(k in 列标签 for k in ("金额", "余额", "收入", "支出", "差额")):
            样式.数字格式 = "#,##0.00;[Red]-#,##0.00"
        if isinstance(值, (_dt, _date)):
            样式.数字格式 = "yyyy-mm-dd"
        if 样式.超链接 or (isinstance(值, str) and 值.startswith("=HYPERLINK(")):
            样式.字色 = "0563C1"
            样式.下划线 = "single"
            样式.加粗 = False
        return 样式
    return 变换


def _紧凑样式(tp: 表计划, *, reading):
    """翻译 紧凑报告.format_compact_sheet。"""
    hs = _列定位(tp)
    df = tp.表
    hidden = {"事项编号"} if tp.名称 in ("待处理事项", "对应明细") else set()
    if reading and tp.名称 in ("每日统计", "月度核对"):
        hidden.update(l for l in hs
                      if "笔数" in l or "变动净额" in l
                      or l.startswith(("已确认", "未确认")))
    first = min(hs.values(), default=1)
    while len(tp.前部列) < first - 1:
        tp.前部列.append(列计划())
    for column in range(1, first):
        tp.前部列[column - 1].隐藏 = True
    for label, column in hs.items():
        cp = tp.列计划取(column - tp.前置空列 - 1)
        cp.隐藏 = label in hidden
        amount = any(k in label for k in ("金额", "余额", "收入", "支出", "差额")) \
            and not any(k in label for k in ("说明", "比例", "依据", "原因"))
        cp.宽 = _紧凑列宽.get(
            label, 14 if amount
            else 28 if any(k in label for k in ("摘要", "原因", "说明", "依据", "记录"))
            else 16)
    last = max(hs.values(), default=1)
    for j in range(last - tp.前置空列, len(tp.列)):
        tp.列[j].隐藏 = True
    tp.tab色 = _紧凑阅读色 if reading else _紧凑留痕色
    tp.状态 = "visible"
    tp.网格线 = False
    tp.缩放 = 100
    tp.左上角 = f"{_字母(first)}1"
    tp.冻结 = None if tp.名称 == "核对总览" else "A2"
    tp.默认行高 = 15
    tp.行计划取(1).高 = 28
    nrows = len(df) + 1
    tp.筛选 = None if tp.名称 == "核对总览" else f"{_字母(first)}1:{_字母(last)}{nrows}"

    # 表头：紧凑样式整体替换字体/填充/对齐，边框保留；A 列等前置空列表头
    # 旧代码因无值而跳过，这里把旧表头样式冻结进覆盖，保持原样。
    旧头 = tp.表头样式
    if tp.表头含前列 and 旧头 is not None:
        for k in range(tp.前置空列):
            tp.覆盖[(1, k + 1)] = 合并样式(_copy(旧头), tp.覆盖.get((1, k + 1)))
    h = 旧头 or 格样式()
    h.字体名, h.字号, h.加粗, h.字色 = "MiSans", 9, True, "243746"
    h.填充色, h.填充类型 = tp.tab色, "solid"
    h.水平, h.垂直, h.自动换行 = "", "center", True
    tp.表头样式 = h

    # 数据行：金额列宽按内容扩、阅读表逐字宽算行高；非阅读表行高抹成 15
    # （快速通道大表只扫前 2000 行定列宽，不再逐行建行计划，行高走表级默认）
    大表 = len(df) > _快速通道行数
    扫描 = 2000 if 大表 else len(df)
    for ri in range(扫描):
        row = ri + 2
        height = 15
        for label, column in hs.items():
            v = _读值(tp, row, column)
            if v is None:
                continue
            cp = tp.列计划取(column - tp.前置空列 - 1)
            if isinstance(v, (int, float, _Decimal)) and not isinstance(v, bool) \
                    and any(k in label for k in ("金额", "余额", "收入", "支出", "差额")):
                cp.宽 = max(cp.宽 or 0, len(f"{v:,.2f}") + 1)
            if reading and label not in hidden \
                    and not (isinstance(v, str) and v.startswith("=")):
                width = cp.宽 or 0
                tw = sum(1.5 if _uni.east_asian_width(ch) in ("W", "F") else 0.75
                         for ch in str(v))
                height = max(height, _math.ceil(tw / max(1, width - 1)) * 12 + 3)
        rp = tp.行计划取(row)
        rp.高 = height
        if not reading:
            rp.隐藏 = False
            rp.收起 = False
    if 大表 and not reading:
        for rp in tp.行.values():
            rp.隐藏 = False
            rp.收起 = False
    tp.格变换 = _做变换(reading)

    p = dict(tp.打印)
    if tp.名称 != "核对总览":
        p["title_rows"] = "1:1"
    p["area"] = f"{_字母(first)}1:{_字母(last)}{nrows}"
    if reading:
        p.update(orientation="landscape", paperSize=9, fitToWidth=1,
                 fitToHeight=0, fitToPage=True)
    tp.打印 = p


def 应用紧凑层(簿: 簿计划):
    """翻译 紧凑报告.apply_compact_report，必须在复核层/展示层回放之后调用。"""
    _写事项(簿)
    _期间标记(簿, "每日统计")
    _期间标记(簿, "月度核对")
    _建总览(簿)
    for tp in 簿.表:
        _紧凑样式(tp, reading=tp.名称 in _紧凑阅读表)
    待处理 = 簿.取表("待处理事项")
    for v in 待处理.验证:
        for row in _sqref行(v.get("sqref", "")):
            待处理.后覆盖[(row, 8)] = 格样式(填充色="FFF2CC", 填充类型="solid")
    for name in ("待处理事项", "对应明细"):
        t = 簿.取表(name)
        hs = _列定位(t)
        for label in ("收入差额", "支出差额"):
            column = _字母(hs[label])
            if len(t.表) > 0:
                t.条件格式.append((
                    f"{column}2:{column}{len(t.表) + 1}",
                    条件规则(公式=[f"AND(ISNUMBER({column}2),{column}2<>0)"],
                            填充色="FFF2CC", 字色="9C0006")))
    for name in ("每日统计", "月度核对"):
        t = 簿.取表(name)
        column = _字母(_列定位(t)["需查看"])
        if len(t.表) > 0:
            t.条件格式.append((
                f"{column}2:{column}{len(t.表) + 1}",
                条件规则(公式=[f'OR({column}2="有差额",{column}2="余额不连续")'],
                         填充色="FFF2CC", 字色="9C0006")))
    order = [簿.取表(n) for n in _紧凑阅读表]
    簿.表 = order + [t for t in 簿.表 if all(t is not o for o in order)]
    簿.active = 0


# ═══════════════════════════════════════════════════════════════════════════
# 总装：按旧路径同序回放全部排版层（reporter._generate_report 的排版段）
# ═══════════════════════════════════════════════════════════════════════════

def 生成簿计划(reporter, tables: dict, effective_config=None, theme="deep-navy") -> 簿计划:
    """tables 为 build_report_tables + apply_column_preferences 之后的表集合。

    回放次序与旧路径一致：底层主题 → 期初警告 → 明细着色 → 差额着色 →
    通用排版 → 复核层 → 展示层 → 期初警告(第二次) → 紧凑层。
    """
    簿 = 新建簿计划(tables)
    for tp in 簿.表:
        应用底层主题(tp, theme)
    has_warning = bool(
        getattr(reporter, "initial_balance_warning", None)
        and reporter.initial_balance_warning.has_warning)
    结论 = 簿.取表("核对结论")
    if 结论 is not None:
        应用期初警告(结论, has_warning)
    for name in ("逐笔匹配", "整组勾稽", "疑点事项", "自动归集事项"):
        t = 簿.取表(name)
        if t is not None:
            应用明细着色(t)
    for name in ("每日统计", "月度统计", "余额差异明细", "余额连续性异常"):
        t = 簿.取表(name)
        if t is not None:
            应用差额着色(t)
    应用通用排版(簿)
    应用复核层(簿)
    应用展示层(簿)
    if 结论 is not None:
        应用期初警告(结论, has_warning)
    应用紧凑层(簿)
    return 簿
