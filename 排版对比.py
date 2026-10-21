# -*- coding: utf-8 -*-
"""逐格对比两份 Excel 工作簿，供排版新旧路径的一致性验收。

读取引擎为 直读XML（纯标准库），不再依赖 openpyxl；仅支持 xlsx。
两侧文件用同一读取器取快照后逐项比对：值、样式、行列尺寸、
冻结/筛选/打印/合并/数据验证/条件格式、计算设置、活动表。

用法:
    python 排版对比.py 旧文件.xlsx 新文件.xlsx
    from 排版对比 import 对比工作簿
    差异 = 对比工作簿(旧, 新)   # 空列表 = 完全一致
"""
import sys
from datetime import date, datetime

from 直读XML import 列转字母, 读簿

容差 = 1e-9


def _色规范(色):
    """颜色归一：8 位 rgb 去 alpha 前缀（00/FF → 后 6 位大写）；
    theme:N / indexed:N 原样保留。"""
    if 色 is None:
        return None
    色 = str(色)
    if 色.startswith(("theme:", "indexed:")):
        return 色
    if len(色) == 8 and 色[:2].upper() in ("00", "FF"):
        return 色[2:].upper()
    return 色.upper()


def _值(value):
    if isinstance(value, float):
        return ("数", round(value, 9))
    if isinstance(value, (datetime, date)):
        return ("日", value.isoformat())
    if value is None:
        return None
    return ("文", str(value))


def _类型(t, 行, 列, 值):
    """对齐 openpyxl data_type：f 公式 / b 布尔 / d 日期 / s 文本 / n 数字。"""
    if (行, 列) in t.公式:
        return "f"
    if 值 is None:
        return None
    if isinstance(值, bool):
        return "b"
    if isinstance(值, (datetime, date)):
        return "d"
    if isinstance(值, str):
        return "s"
    return "n"


def _边规范(边):
    if 边 is None:
        return None
    return (边[0], _色规范(边[1]))


def _格快照(t, 行, 列, 库):
    样式 = t.样式(行, 列, 库)
    字体 = 样式["字体"] or (None, None, False, None, None)
    填充 = 样式["填充"] or (None, None)
    边框 = 样式["边框"] or (None, None, None, None)
    对齐 = 样式["对齐"]
    格式 = 样式["数字格式"]
    值 = t.格.get((行, 列))
    return {
        "值": _值(值),
        "类型": _类型(t, 行, 列, 值),
        "字体": (字体[0], 字体[1], bool(字体[2]), _色规范(字体[3]), 字体[4]),
        "填充": (填充[0], _色规范(填充[1])),
        "对齐": (对齐[0], 对齐[1], bool(对齐[2])),
        "边框": tuple(_边规范(b) for b in 边框),
        "数字格式": None if 格式 in (None, "General") else str(格式),
        "批注": t.批注.get((行, 列)),
        "超链接": t.超链接.get((行, 列)),
    }


_格属性中文 = {"值": "值", "类型": "类型", "字体": "字体", "填充": "填充", "对齐": "对齐",
             "边框": "边框", "数字格式": "数字格式", "批注": "批注", "超链接": "超链接"}


def _验证快照(v):
    """v 为读取器元组 (sqref, type, allowBlank, formula1, formula2,
    errorStyle, errorTitle, error, showErrorMessage)。"""
    sqref, 型, 空, f1, f2, 错式, 错题, 错文, 报错 = v
    return (型, f1, f2, bool(空), 错文, 错题, bool(报错), 错式,
            tuple(sorted((sqref or "").split())))


def _条件格式快照(条目们):
    """条目 (sqref, type, 公式们, dxf描述)；dxf描述 为 (字色, 填充色) 或 None。"""
    结果 = []
    for sqref, 型, 公式们, dxf in 条目们:
        字色, 填充 = dxf if dxf else (None, None)
        结果.append((sqref, 型, tuple(公式们 or ()), _色规范(填充), _色规范(字色)))
    return sorted(结果, key=repr)


def _表级快照(t):
    打印 = t.打印
    设置 = 打印.get("setup", {})
    边距 = 打印.get("margins", {})
    return {
        "状态": t.状态,
        "tab色": _色规范(t.tab色),
        "冻结": t.冻结,
        "筛选": t.筛选,
        "网格线": bool(t.网格线),
        "缩放": t.缩放 or 100,
        "outline汇总在下": t.outline汇总在下,
        "打印": (设置.get("orientation"), 设置.get("paperSize"),
               设置.get("fitToWidth"), 设置.get("fitToHeight"),
               打印.get("fitToPage"),
               打印.get("title_rows"), 打印.get("area"),
               round(float(边距["left"]), 4) if 边距.get("left") else None,
               round(float(边距["right"]), 4) if 边距.get("right") else None),
        "合并": tuple(sorted(t.合并)),
        "数据验证": sorted((_验证快照(v) for v in t.数据验证), key=repr),
        "条件格式": _条件格式快照(t.条件格式),
    }


def _快照(簿):
    库 = 簿.样式库
    数据 = {}
    for t in 簿:
        格子 = {}
        # 值格 + 公式格 + 显式非 0 样式格（s=0 无值格与 openpyxl has_style 同义跳过）
        集合 = set(t.格) | set(t.公式) | {k for k, s in t.格样式索引.items() if s}
        for (行, 列) in 集合:
            格子[(行, 列)] = _格快照(t, 行, 列, 库)
        行 = {号: {"行高": d["高"], "隐藏": bool(d["隐藏"]),
                 "分级": d["分级"], "收起": bool(d["收起"])}
            for 号, d in t.行.items()}
        列 = {列转字母(号): {"列宽": d["宽"], "隐藏": bool(d["隐藏"])}
            for 号, d in t.列.items()}
        数据[t.名称] = {"格": 格子, "行": 行, "列": 列, "表级": _表级快照(t),
                     "默认行高": t.默认行高}
    return 数据


def _计算快照(设置):
    """calcPr 归一：calcMode 缺省 auto，两个开关缺省 False。"""
    return {
        "calcMode": 设置.get("calcMode", "auto"),
        "fullCalcOnLoad": 设置.get("fullCalcOnLoad") in ("1", "true"),
        "forceFullCalc": 设置.get("forceFullCalc") in ("1", "true"),
    }


def _比字典(差异, 位置, 旧, 新, 键名):
    for 键 in sorted(set(旧) | set(新)):
        旧值, 新值 = 旧.get(键), 新.get(键)
        if 旧值 != 新值:
            差异.append(f"{位置}: {键名}{键} 旧={旧值} 新={新值}")


def 对比工作簿(旧路径, 新路径, 差异上限=200):
    """逐格对比两份工作簿，返回差异描述列表；空列表表示完全一致。"""
    旧簿 = 读簿(旧路径)
    新簿 = 读簿(新路径)
    差异 = []

    if 旧簿.表名列表 != 新簿.表名列表:
        差异.append(f"工作表顺序或名单: 旧={旧簿.表名列表} 新={新簿.表名列表}")

    旧计算 = _计算快照(旧簿.计算设置)
    新计算 = _计算快照(新簿.计算设置)
    for 属性 in ("calcMode", "fullCalcOnLoad", "forceFullCalc"):
        if 旧计算[属性] != 新计算[属性]:
            差异.append(f"工作簿计算设置 {属性}: 旧={旧计算[属性]} 新={新计算[属性]}")
    if 旧簿.active名 != 新簿.active名:
        差异.append(f"活动表: 旧={旧簿.active名} 新={新簿.active名}")

    旧数据 = _快照(旧簿)
    新数据 = _快照(新簿)
    for 表名 in 旧数据.keys() & 新数据.keys():
        旧表, 新表 = 旧数据[表名], 新数据[表名]
        _比字典(差异, f"{表名}", 旧表["表级"], 新表["表级"], "")
        _比字典(差异, f"{表名}", 旧表["列"], 新表["列"], "列")
        for 行号 in sorted(set(旧表["行"]) | set(新表["行"])):
            旧行 = 旧表["行"].get(行号, {"行高": None, "隐藏": False, "分级": 0, "收起": False})
            新行 = 新表["行"].get(行号, {"隐藏": False, "分级": 0, "收起": False, "行高": None})
            旧缺省 = 旧表["默认行高"] or "默认"
            新缺省 = 新表["默认行高"] or "默认"
            for 键 in ("隐藏", "分级", "收起"):
                if 旧行[键] != 新行[键]:
                    差异.append(f"{表名} 第{行号}行: {键} 旧={旧行[键]} 新={新行[键]}")
            旧高 = 旧行["行高"] if 旧行["行高"] is not None else 旧缺省
            新高 = 新行["行高"] if 新行["行高"] is not None else 新缺省
            if 旧高 != 新高:
                差异.append(f"{表名} 第{行号}行: 行高 旧={旧行['行高']} 新={新行['行高']}")
        for (行号, 列号) in sorted(set(旧表["格"]) | set(新表["格"])):
            旧格 = 旧表["格"].get((行号, 列号))
            新格 = 新表["格"].get((行号, 列号))
            坐标 = f"{列转字母(列号)}{行号}"
            if 旧格 is None or 新格 is None:
                缺失 = 旧格 if 旧格 is not None else 新格
                if 缺失 and (缺失["值"] is not None or any(v is not None for k, v in 缺失.items() if k != "值")):
                    差异.append(f"{表名}!{坐标}: 格缺失 旧={'有' if 旧格 else '无'} 新={'有' if 新格 else '无'}")
                continue
            for 键, 中文 in _格属性中文.items():
                if 旧格[键] != 新格[键]:
                    差异.append(f"{表名}!{坐标}: {中文} 旧={旧格[键]} 新={新格[键]}")
            if len(差异) >= 差异上限:
                差异.append(f"……差异超过 {差异上限} 条，截断")
                return 差异
    return 差异


def 对比并报告(旧路径, 新路径, 标签=""):
    差异 = 对比工作簿(旧路径, 新路径)
    前缀 = f"[{标签}] " if 标签 else ""
    if 差异:
        print(f"{前缀}不一致，共 {len(差异)} 条差异：")
        for 条目 in 差异[:30]:
            print("  " + 条目)
        if len(差异) > 30:
            print(f"  ……其余 {len(差异) - 30} 条省略")
        return False
    print(f"{前缀}逐格一致")
    return True


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("用法: python 排版对比.py 旧文件.xlsx 新文件.xlsx")
        sys.exit(2)
    sys.exit(0 if 对比并报告(sys.argv[1], sys.argv[2]) else 1)
