#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_excel — 从 DataFrame 生成摩根士丹利标准格式的 Excel 报表。
beautify — 对已有 Excel 文件只改格式不改数据（值与公式原文保留；公式缓存值
不由本工具计算，文件带自动重算标记，Excel 打开时重算）。

用法:
    from make_excel import make_excel, beautify
    make_excel(df, 'output.xlsx')
    beautify('input.xlsx')              # 原地美化
    beautify('input.xlsx', 'out.xlsx')  # 另存美化

    python make_excel.py data.csv output.xlsx
    python make_excel.py --beautify input.xlsx output.xlsx

引擎: 直写XML（写）+ 直读XML（读），不再依赖 openpyxl；仅支持 xlsx。
设计: DataFrame → 排版计划.表计划 → 直写XML 一次 zip 成型。
限制: beautify 要求表头在第 1 行（上方有空行/标题行的文件会明确报错）。
"""

import argparse
import re
from datetime import datetime, date
from pathlib import Path
from shutil import copy2
from typing import Optional

import pandas as pd

from 排版共享 import (
    A_COL_WIDTH, FMT, MAX_COL_WIDTH, MIN_COL_WIDTH, ROW_HEIGHT, SUM_KEYWORDS,
    THEMES, atomic_output_path,
    列类型推断, 列宽估算,
)
from 排版计划 import 表计划, 簿计划, 格样式, 合并样式, 列计划, 应用底层主题
from 直写XML import 写簿

__all__ = [
    "make_excel", "beautify", "atomic_output_path",
    "THEMES", "FMT", "SUM_KEYWORDS", "ROW_HEIGHT", "A_COL_WIDTH",
    "MIN_COL_WIDTH", "MAX_COL_WIDTH",
]


# 兼容别名：测试与外部调用沿用旧名
_infer_column_type = 列类型推断
_estimate_column_width = 列宽估算


def _char_width(text) -> int:
    """中文字符计2，拉丁/数字计1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))

AVAILABLE_THEMES = ", ".join(sorted(THEMES.keys()))


def _validate_theme(theme_name: str):
    """校验主题名，无效则抛 ValueError。"""
    if theme_name not in THEMES:
        raise ValueError(
            f"未知主题 '{theme_name}'。可用主题: {sorted(THEMES.keys())}"
        )


# ═══════════════════════════════════════════════════════════════════════════
# 主入口
# ═══════════════════════════════════════════════════════════════════════════

def make_excel(data, output_path: str, sheet_name: str = 'Sheet1', theme: str = 'default',
               fmt_override: Optional[dict] = None, 首列留白: bool = True) -> str:
    """生成摩根系标准格式 Excel。

    参数
    ----
    data : DataFrame | list[tuple[str, DataFrame]]
        单 sheet 传 df，多 sheet 传 [(名称, df), ...]。
    output_path : str
        输出文件路径。
    sheet_name : str
        单 sheet 模式的工作表名称。
    theme : str
        色系主题名。可用: """ + AVAILABLE_THEMES + """。默认 'default'。
    fmt_override : dict, optional
        自定义数字格式，如 {'money': '#,##0.0', 'pct': '0.0%'}。
        合并到默认 FMT 之上，同名覆盖。
    首列留白 : bool, default True
        True 时 A 列留 2 宽空列、数据从 B 列开始（审计报告版式）；
        False 时数据从 A 列开始（导入模板等直接填写的场景）。

    返回
    ----
    str : 输出文件的绝对路径。
    """
    _validate_theme(theme)

    # 统一为多 sheet 格式
    if isinstance(data, pd.DataFrame):
        sheets = [(sheet_name, data)]
    else:
        sheets = list(data)

    计划 = 簿计划()
    for name, df in sheets:
        if df is None:
            continue
        safe_name = re.sub(r'[\\/*?:\[\]]', '_', str(name))[:31]
        tp = 表计划(名称=safe_name, 表=df)
        应用底层主题(tp, theme=theme, fmt_override=fmt_override, 首列留白=首列留白)
        计划.表.append(tp)
    if not 计划.表:
        raise ValueError("没有可写入的工作表（全部数据为空）")

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    写簿(计划, str(out))
    return str(out.absolute())


# ═══════════════════════════════════════════════════════════════════════════
# beautify — 只改格式不改数据
# ═══════════════════════════════════════════════════════════════════════════

def _检测数据区(t) -> tuple[int, int, int] | None:
    """从 表读 检测数据区。返回 (表头行, 数据末行, 末列)；空表返回 None。

    规则：第一个非空行为表头行；最后一个非空行为数据末行；
    末列取全表非空格最大列。表头必须在第 1 行，否则明确报错。
    """
    if not t.格:
        return None
    非空行 = sorted({r for r, _c in t.格})
    表头行 = 非空行[0]
    if 表头行 != 1:
        raise ValueError(
            f"工作表“{t.名称}”的表头不在第 1 行（在第 {表头行} 行），"
            "beautify 暂不支持首行以上有空行或标题行的文件"
        )
    末列 = max(c for _r, c in t.格)
    return 表头行, 非空行[-1], 末列


def _列类型自值(表头, 值列: list) -> str:
    """beautify 的列类型推断：关键词规则与 列类型推断 一致，值特征直接看列表。"""
    name = str(表头) if 表头 is not None else ''
    n = name.lower()
    if re.search(r'(率$|占比|百分比|增长率|利润率|达成率|rate$|ratio$|margin|pct|percent)', n):
        return 'pct'
    if re.search(r'(日期|时间|年月|期间|年份|月份|date|time|year|month)', n):
        return 'date'
    if re.search(r'(编号|id|单号|编码|电话|手机|备注|说明|名称|地址|描述|号码|订单号|负责人|姓名|联系人|部门|岗位|phone|email|code|desc|address)', n):
        return 'text'
    if re.search(r'(金额|余额|价格|收入|成本|费用|毛利额|净利|合计|总额|售价|单价|预算|支出|毛利(?!率)|返点|佣金|折扣|补贴|周转|本金|amount|balance|price|cost|revenue|total|budget|expense|commission|discount|rebate|profit|spending)', n):
        return 'money'

    样本 = [v for v in 值列 if v is not None][:100]
    if not 样本:
        return 'unknown'
    if sum(1 for v in 样本 if isinstance(v, (datetime, date))) >= len(样本) * 0.5:
        return 'date'
    数值 = [v for v in 样本 if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if 数值:
        浮点 = [v for v in 数值 if isinstance(v, float)]
        if 浮点:
            if all(0 < v < 1 for v in 浮点) and len(set(浮点)) > 2:
                return 'pct'
            return 'money'
        return 'number'
    if any(re.match(r'^\d{12,}$', str(v)) for v in 样本 if not isinstance(v, (int, float))):
        return 'text'
    return 'unknown'


def _重建表计划(t, theme: str, fmt_override: Optional[dict],
                col_types: Optional[dict]) -> 表计划:
    """把 表读 重建为套好摩根系格式的 表计划（值、公式、合并、批注、超链接保留）。"""
    主题 = THEMES[theme]
    有效格式 = dict(FMT)
    if fmt_override:
        有效格式.update(fmt_override)

    检测 = _检测数据区(t)
    # 全内容重建：第 1 行作表头，其余作数据行；无内容的表建空 df。
    # 列范围只取有内容的列（纯宽度定义的尾列不重建，避免空表头被落成空串格）
    末列 = 检测[2] if 检测 else 0
    末行 = 检测[1] if 检测 else 0
    # A 列完全无内容才作留白；df 只从实际起始列建，避免留白+空 A 列双重右移
    a有数据 = any(t.格.get((r, 1)) is not None for r in range(1, 末行 + 1))
    起建列 = 1 if a有数据 else 2
    if 末列 < 起建列:
        末列 = 起建列 - 1 if a有数据 else 0
    表头 = {c: t.格.get((1, c)) for c in range(起建列, 末列 + 1)}
    列名 = [(str(v) if v is not None else "") for v in 表头.values()]
    # 列名去重（pandas 允许重名，但 dict 化场景会乱；追加序号保持唯一）
    见过 = {}
    for i, 名 in enumerate(列名):
        if 名 in 见过:
            见过[名] += 1
            列名[i] = f"{名}.{见过[名]}"
        else:
            见过[名] = 0
    行数据 = [[t.格.get((r, c)) for c in range(起建列, 末列 + 1)]
              for r in range(2, 末行 + 1)]
    df = pd.DataFrame(行数据, columns=列名, dtype=object)
    tp = 表计划(名称=t.名称, 表=df)
    tp.合并 = set(t.合并)
    tp.tab色 = t.tab色
    tp.状态 = t.状态

    if 检测 is None:
        return tp
    表头行, 数据末行, 数据末列 = 检测

    # A 列：有数据则并入数据区（前置空列=0），否则作 2 宽留白
    if a有数据:
        tp.前置空列 = 0
        起列 = 1
    else:
        tp.前置空列 = 1
        tp.前部列 = [列计划(宽=A_COL_WIDTH)]
        起列 = 2

    # 行高
    for r in range(1, 数据末行 + 1):
        tp.行计划取(r).高 = ROW_HEIGHT

    # 表头样式（第 1 行，含 A 列有数据时的 A1）
    tp.表头样式 = 格样式(
        字体名="Arial", 字号=11, 加粗=True, 字色=主题["header_font_color"],
        填充色=主题["header_fill"], 填充类型="solid",
        水平="right", 垂直="center",
        上框=("medium", "000000"), 下框=("dashed", "808080"),
    )

    # 数据列：类型推断（支持手动覆盖）+ 列宽 + 列基础样式（含中行虚线边框）
    中虚 = ("dashed", "808080")
    for c in range(起列, 数据末列 + 1):
        j = c - tp.前置空列 - 1
        名 = str(表头[c]) if 表头.get(c) is not None else ""
        if col_types and 名 in col_types:
            ct = col_types[名]
        else:
            ct = _列类型自值(表头.get(c), df.iloc[:, j].tolist())
        cp = tp.列计划取(j)
        w = _char_width(名)
        for v in df.iloc[:, j].tolist():
            if v is None:
                continue
            try:
                w = max(w, _char_width(str(v)))
            except Exception:
                pass
        cp.宽 = max(MIN_COL_WIDTH, min(w + 2, MAX_COL_WIDTH))
        数字列 = ct in ("money", "number", "pct")
        cp.基础样式 = 格样式(
            字体名="Arial", 字号=11,
            字色=主题["data_font_blue"] if 数字列 else None,
            水平="right" if 数字列 else "left", 垂直="center",
            数字格式=有效格式.get(ct) or None,
            上框=中虚, 下框=中虚,
        )
        # 逐格修正：公式格黑字；date 列里的原始数字回落 General
        for ri in range(len(df)):
            r = ri + 2
            if (r, c) in t.公式 and 数字列:
                tp.盖格(r, c, 格样式(字色=主题["data_font_black"]))
            if ct == "date":
                v = df.iat[ri, j]
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    tp.盖格(r, c, 格样式(数字格式="General"))

    # 合计行：末行首数据列（或 A 列）含合计关键词 → 数据区加粗 + 填充
    首值 = t.格.get((数据末行, 1)) if a有数据 else None
    if not 首值 or not any(kw in str(首值) for kw in SUM_KEYWORDS):
        首值 = t.格.get((数据末行, 起列))
    if 数据末行 > 表头行 and 首值 is not None \
            and any(kw in str(首值) for kw in SUM_KEYWORDS):
        for c in range(起列, 数据末列 + 1):
            tp.盖格(数据末行, c, 格样式(
                字体名="Arial", 字号=11, 加粗=主题["sum_font_bold"],
                字色=主题["sum_font_color"],
                填充色=主题["sum_fill"], 填充类型="solid"))

    # 末行下框加粗
    if 数据末行 > 1:
        rp = tp.行计划取(数据末行)
        rp.样式 = 合并样式(rp.样式, 格样式(上框=中虚, 下框=("medium", "000000")))

    # 右侧空白列（宽 3）
    tp.列计划取(数据末列 - tp.前置空列).宽 = 3

    # 批注 / 超链接保留（随样式覆盖走）
    for (r, c), (文本, 作者) in t.批注.items():
        tp.盖格(r, c, 格样式(批注=(文本, 作者)))
    for (r, c), 目标 in t.超链接.items():
        tp.盖格(r, c, 格样式(超链接=目标))

    tp.网格线 = False
    tp.冻结 = f"A{表头行 + 1}"
    return tp


def beautify(
    input_path: str,
    output_path: Optional[str] = None,
    *,
    col_types: Optional[dict] = None,
    backup: bool = True,
    theme: str = 'default',
    fmt_override: Optional[dict] = None,
) -> str:
    """美化已有 Excel 文件，只改格式不改数据（保留公式原文和值）。

    公式缓存值不由本工具计算：输出文件带自动重算标记，Excel 打开时重算。
    要求各表表头在第 1 行；不满足时抛 ValueError，不静默改动布局。

    参数
    ----
    input_path : str
        输入的 xlsx 文件路径。
    output_path : str, optional
        输出路径。不传则原地覆盖原文件。
    col_types : dict, optional
        手动指定列类型，如 {'订单号': 'text', '金额': 'money'}。
        覆盖自动类型推断。可选值: money/number/pct/date/text。
    backup : bool, default True
        原地覆盖前是否自动备份。备份文件名为 <原文件名>_backup_YYYYMMDD_HHMMSS.xlsx。
    theme : str
        色系主题名。可用: """ + AVAILABLE_THEMES + """。默认 'default'。
    fmt_override : dict, optional
        自定义数字格式，如 {'money': '#,##0.0', 'pct': '0.0%'}。
        合并到默认 FMT 之上，同名覆盖。

    返回
    ----
    str : 输出文件的绝对路径。
    """
    _validate_theme(theme)
    from 直读XML import 读簿

    out = output_path or input_path

    # 自动备份
    if backup and output_path is None:
        inp = Path(input_path)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        bak = inp.with_stem(f'{inp.stem}_backup_{ts}')
        copy2(input_path, str(bak))
        print(f'备份: {bak}')

    簿 = 读簿(input_path)   # 公式格取原文（"=..."），重建后仍是活公式

    计划 = 簿计划()
    for t in 簿:
        计划.表.append(_重建表计划(t, theme, fmt_override, col_types))
    if 簿.active名 in 簿.表名列表:
        计划.active = 簿.表名列表.index(簿.active名)

    out_path = Path(out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    写簿(计划, str(out_path))
    return str(out_path.absolute())


def main():
    parser = argparse.ArgumentParser(description='生成/美化摩根系标准格式 Excel')
    parser.add_argument('input', help='输入文件（CSV 或 --beautify 时传 xlsx）')
    parser.add_argument('output', nargs='?', default=None,
                        help='输出路径（默认: 桌面/<输入文件名>.xlsx）')
    parser.add_argument('--sheet-name', default='Sheet1', help='工作表名称（仅 CSV 模式）')
    parser.add_argument('--beautify', action='store_true',
                        help='美化已有 xlsx（只改格式不改数据，保留公式）')
    parser.add_argument('--theme', default='default', choices=list(THEMES.keys()),
                        help=f"色系主题 (default: default)。可选: {AVAILABLE_THEMES}")
    args = parser.parse_args()

    if args.beautify:
        path = beautify(args.input, args.output, theme=args.theme)
        print(f'已美化: {path}')
        return

    if args.output is None:
        in_path = Path(args.input)
        args.output = str(Path.home() / 'Desktop' / in_path.with_suffix('.xlsx').name)

    df = pd.read_csv(args.input)
    path = make_excel(df, args.output, sheet_name=args.sheet_name, theme=args.theme)
    print(f'已生成: {path}')


if __name__ == '__main__':
    main()
