# -*- coding: utf-8 -*-
"""openpyxl 假面（仅测试用）：读走 直读XML，写走 排版计划+直写XML。

只实现测试用到的最小 API 子集，不是 openpyxl 的完整替身。
生产代码禁止 import 本模块。
"""
from collections import defaultdict
from datetime import date, datetime
from types import SimpleNamespace

import pandas as pd

from 直读XML import 读簿, 拆坐标, 列转字母, 字母转列
from 排版共享 import 条件规则
from 排版计划 import 簿计划, 表计划, 格样式
from 直写XML import 写簿

get_column_letter = 列转字母
column_index_from_string = 字母转列


def 存表(df, 路径, *, header=True, 表名="Sheet1"):
    """df.to_excel(index=False) 的测试替身：排版计划+直写XML 落盘。

    header=False 时第一行数据落为文件第 1 行（表头格即普通值格）。
    """
    def _原(v):
        if v is not None and type(v).__module__ == "numpy":
            v = v.item()
        return v

    母 = df.astype(object).where(pd.notna(df), None)
    if header:
        列名 = [str(c) for c in 母.columns]
        体 = 母
    else:
        首行 = 母.iloc[0]
        列名 = ["" if v is None else str(_原(v)) for v in 首行]
        体 = 母.iloc[1:]
    数据 = [[_原(v) for v in 行] for 行 in 体.itertuples(index=False, name=None)]
    tp = 表计划(名称=表名, 表=pd.DataFrame(数据, columns=列名))
    写簿(簿计划(表=[tp]), str(路径))


def 读Excel(路径, sheet_name=0):
    """pd.read_excel 的测试替身：第 1 行作表头，数据从第 2 行起。

    空格列表头按 pandas 惯例叫 Unnamed: N；重名表头追加 .1/.2。
    sheet_name=None 返回 {表名: DataFrame}。
    """
    簿 = 读簿(路径, 数据=True, 读样式=False)

    def _df(名):
        t = 簿.表[名]
        行 = list(t.行值(起=1, 止=t.最大行))
        if not 行:
            return pd.DataFrame()
        列名 = []
        见过 = {}
        for i, v in enumerate(行[0]):
            名 = str(v) if v is not None else f"Unnamed: {i}"
            if 名 in 见过:
                见过[名] += 1
                名 = f"{名}.{见过[名]}"
            else:
                见过[名] = 0
            列名.append(名)
        return pd.DataFrame(行[1:], columns=列名)

    if sheet_name is None:
        return {名: _df(名) for 名 in 簿.表名列表}
    if isinstance(sheet_name, int):
        sheet_name = 簿.表名列表[sheet_name]
    return _df(sheet_name)


# ═══════════════════════════════════════════════════════════════════════════
# 读侧
# ═══════════════════════════════════════════════════════════════════════════

class _字体:
    def __init__(self, t):
        名, 号, 粗, 色, 下划 = t or (None, None, False, None, None)
        self.name = 名
        self.size = self.sz = 号
        self.bold = self.b = 粗
        self.color = _颜色壳(色) if 色 else None
        self.underline = 下划


class _批注:
    def __init__(self, 文本, 作者):
        self.text, self.author = 文本, 作者


class _填充壳:
    def __init__(self, 类型, 色):
        类型 = None if 类型 in (None, "none") else 类型
        self.fill_type = self.patternType = 类型
        self.fgColor = self.start_color = _颜色壳(色) if 色 else None


class _边壳:
    def __init__(self, 边):
        if 边:
            self.style = 边[0]
            self.color = _颜色壳(边[1]) if 边[1] else None
        else:
            self.style = None
            self.color = None


class _边框壳:
    def __init__(self, 边框):
        上, 下, 左, 右 = 边框 or (None, None, None, None)
        self.top, self.bottom, self.left, self.right = (_边壳(b) for b in (上, 下, 左, 右))


class _对齐壳:
    def __init__(self, 对齐):
        水平, 垂直, 换行 = 对齐 or (None, None, False)
        self.horizontal, self.vertical, self.wrap_text = 水平, 垂直, 换行


class _超链:
    def __init__(self, 目标):
        self.target = 目标
        self.location = None


class 格壳:
    __slots__ = ("_t", "_库", "row", "column", "_数")

    def __init__(self, t, 库, 行, 列, 数):
        self._t, self._库, self._数 = t, 库, 数
        self.row, self.column = 行, 列

    @property
    def value(self):
        return self._t.格.get((self.row, self.column))

    @property
    def coordinate(self):
        return f"{列转字母(self.column)}{self.row}"

    @property
    def column_letter(self):
        return 列转字母(self.column)

    @property
    def data_type(self):
        if (self.row, self.column) in self._t.公式 and not self._数:
            return "f"
        v = self.value
        if v is None:
            return "n"
        if isinstance(v, bool):
            return "b"
        if isinstance(v, (datetime, date)):
            return "d"
        if isinstance(v, str):
            return "s"
        return "n"

    @property
    def is_date(self):
        return isinstance(self.value, (datetime, date))

    @property
    def has_style(self):
        return bool(self._t.格样式索引.get((self.row, self.column)))

    def _有效样式(self):
        return self._t.样式(self.row, self.column, self._库)

    @property
    def number_format(self):
        return self._有效样式()["数字格式"]

    @property
    def font(self):
        return _字体(self._有效样式()["字体"])

    @property
    def comment(self):
        c = self._t.批注.get((self.row, self.column))
        return _批注(*c) if c else None

    @property
    def fill(self):
        类型, 色 = self._有效样式()["填充"] or (None, None)
        return _填充壳(类型, 色)

    @property
    def border(self):
        return _边框壳(self._有效样式()["边框"])

    @property
    def alignment(self):
        return _对齐壳(self._有效样式()["对齐"])

    @property
    def hyperlink(self):
        h = self._t.超链接.get((self.row, self.column))
        return _超链(h) if h else None


class _行维:
    __slots__ = ("height", "hidden", "outlineLevel", "collapsed")

    def __init__(self, d):
        d = d or {}
        self.height = d.get("高")
        self.hidden = d.get("隐藏", False)
        self.outlineLevel = d.get("分级", 0)
        self.collapsed = d.get("收起", False)


class _行维度集:
    def __init__(self, t):
        self._t = t

    def __getitem__(self, r):
        return _行维(self._t.行.get(r))


class _列维:
    __slots__ = ("width", "hidden")

    def __init__(self, d):
        d = d or {}
        self.width = d.get("宽")
        self.hidden = d.get("隐藏", False)


class _列维度集:
    def __init__(self, t):
        self._t = t

    def __getitem__(self, 字母):
        return _列维(self._t.列.get(字母转列(str(字母))))


class _区域:
    """sqref 字符串，支持 “B2” in sqref 判断。"""

    def __init__(self, ref):
        self._ref = ref or ""

    def __contains__(self, 坐标):
        坐标 = str(坐标)
        for 片 in self._ref.split():
            if ":" in 片:
                try:
                    r1, c1 = 拆坐标(片.split(":")[0])
                    r2, c2 = 拆坐标(片.split(":")[1])
                    r, c = 拆坐标(坐标)
                except ValueError:
                    continue
                if r1 <= r <= r2 and c1 <= c <= c2:
                    return True
            elif 片 == 坐标:
                return True
        return False

    def __str__(self):
        return self._ref

    def __repr__(self):
        return self._ref

    def __eq__(self, 他):
        return str(self) == str(他)

    def __hash__(self):
        return hash(self._ref)


class _验证:
    def __init__(self, v):
        (sqref, self.type, self.allowBlank, self.formula1, self.formula2,
         self.errorStyle, self.errorTitle, self.error, self.showErrorMessage) = v
        self.sqref = _区域(sqref)
        self.allow_blank = self.allowBlank


class _验证集:
    def __init__(self, t):
        self.dataValidation = [_验证(v) for v in t.数据验证]


class _条件集:
    def __init__(self, t):
        self._n = len(t.条件格式)
        self._cf_rules = {}
        for sq, _类型, 公式们, _dxf in t.条件格式:
            self._cf_rules.setdefault(sq, []).append(公式们)

    def __bool__(self):
        return self._n > 0

    def __len__(self):
        return self._n


class _颜色壳:
    def __init__(self, rgb):
        self.rgb = rgb


class _属性壳:
    def __init__(self, t):
        self.tabColor = _颜色壳(t.tab色) if t.tab色 else None


class _表格式壳:
    def __init__(self, t):
        self.defaultRowHeight = t.默认行高
        self.defaultColWidth = t.默认列宽


class _筛选壳:
    def __init__(self, ref):
        self.ref = ref


class _视图壳:
    def __init__(self, t):
        self.showGridLines = t.网格线
        self.zoomScale = t.缩放
        self.topLeftCell = t.左上角


class _页设壳:
    def __init__(self, t):
        s = t.打印.get("setup", {})
        self.orientation = s.get("orientation")
        self.paperSize = s.get("paperSize")
        self.fitToWidth = int(s["fitToWidth"]) if s.get("fitToWidth") else None
        self.fitToHeight = int(s["fitToHeight"]) if s.get("fitToHeight") else None
        self.fitToPage = t.打印.get("fitToPage")


class _边距壳:
    def __init__(self, t):
        m = t.打印.get("margins", {})
        for k in ("left", "right", "top", "bottom", "header", "footer"):
            setattr(self, k, float(m[k]) if m.get(k) is not None else None)


class 表壳:
    """直读XML.表读 的 openpyxl 形外壳；未实现的属性直通 表读。"""

    def __init__(self, t, 库, 数):
        object.__setattr__(self, "_t", t)
        object.__setattr__(self, "_库", 库)
        object.__setattr__(self, "_数", 数)

    def __getattr__(self, 名):
        # 行值/名称/格/公式/最大行/最大列 等直接透传给 表读
        return getattr(object.__getattribute__(self, "_t"), 名)

    def __eq__(self, 他):
        return isinstance(他, 表壳) and self._t is 他._t

    def __hash__(self):
        return id(self._t)

    @property
    def title(self):
        return self._t.名称

    @property
    def max_row(self):
        return self._t.最大行

    @property
    def max_column(self):
        return self._t.最大列

    @property
    def sheet_state(self):
        return self._t.状态

    @property
    def freeze_panes(self):
        return self._t.冻结

    @property
    def sheet_view(self):
        return _视图壳(self._t)

    @property
    def page_setup(self):
        return _页设壳(self._t)

    @property
    def page_margins(self):
        return _边距壳(self._t)

    @property
    def print_title_rows(self):
        tr = self._t.打印.get("title_rows")
        return "$" + tr.replace(":", ":$") if tr else None

    def cell(self, row, column, value=None):
        if value is not None:
            raise TypeError("读侧假面不支持写值，请改用写侧 Workbook")
        return 格壳(self._t, self._库, row, column, self._数)

    def __getitem__(self, 键):
        if isinstance(键, str):
            行, 列 = 拆坐标(键)
            return 格壳(self._t, self._库, 行, 列, self._数)
        return tuple(格壳(self._t, self._库, 键, c, self._数)
                     for c in range(1, self._t.最大列 + 1))

    def __iter__(self):
        return (self[r] for r in range(1, self._t.最大行 + 1))

    @property
    def values(self):
        return (tuple(g.value for g in 行) for 行 in self)

    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None,
                  values_only=False):
        r1, r2 = min_row or 1, max_row or self._t.最大行
        c1, c2 = min_col or 1, max_col or self._t.最大列
        for r in range(r1, r2 + 1):
            行 = tuple(格壳(self._t, self._库, r, c, self._数)
                      for c in range(c1, c2 + 1))
            if values_only:
                yield tuple(g.value for g in 行)
            else:
                yield 行

    @property
    def sheet_format(self):
        return _表格式壳(self._t)

    @property
    def row_dimensions(self):
        return _行维度集(self._t)

    @property
    def column_dimensions(self):
        return _列维度集(self._t)

    @property
    def data_validations(self):
        return _验证集(self._t)

    @property
    def conditional_formatting(self):
        return _条件集(self._t)

    @property
    def sheet_properties(self):
        return _属性壳(self._t)

    @property
    def auto_filter(self):
        return _筛选壳(self._t.筛选)


class 簿壳:
    def __init__(self, 簿, 数):
        self._簿, self._数 = 簿, 数

    @property
    def calculation(self):
        设置 = self._簿.计算设置
        return SimpleNamespace(
            calcMode=设置.get("calcMode", "auto"),
            fullCalcOnLoad=设置.get("fullCalcOnLoad") in ("1", "true", True),
            forceFullCalc=设置.get("forceFullCalc") in ("1", "true", True),
        )

    @property
    def sheetnames(self):
        return list(self._簿.表名列表)

    def __getitem__(self, 名):
        return 表壳(self._簿.表[名], self._簿.样式库, self._数)

    def __iter__(self):
        return (表壳(t, self._簿.样式库, self._数) for t in self._簿)

    @property
    def active(self):
        return self[self._簿.active名]

    def close(self):
        pass


def load_workbook(path, data_only=False, read_only=False, **_忽略):
    return 簿壳(读簿(path, 数据=data_only), data_only)


# ═══════════════════════════════════════════════════════════════════════════
# 写侧
# ═══════════════════════════════════════════════════════════════════════════

class Font:
    def __init__(self, name=None, size=None, bold=False, color=None, **_kw):
        self.name, self.size, self.bold, self.color = name, size, bold, color


class PatternFill:
    def __init__(self, patternType=None, fill_type=None, fgColor=None, bgColor=None, **_kw):
        self.patternType = fill_type or patternType
        self.fgColor = fgColor
        self.bgColor = bgColor


class Alignment:
    def __init__(self, horizontal=None, vertical=None, wrap_text=None, **_kw):
        self.horizontal, self.vertical, self.wrap_text = horizontal, vertical, wrap_text


class FormulaRule:
    def __init__(self, formula=None, fill=None, font=None, **_kw):
        self.type = "expression"
        self.formula = list(formula or [])
        self.fill = fill
        self.font = font


class DataValidation:
    def __init__(self, type=None, formula1=None, formula2=None, allow_blank=False,
                 errorStyle=None, errorTitle=None, error=None,
                 showErrorMessage=False, **_kw):
        self.type, self.formula1, self.formula2 = type, formula1, formula2
        self.allowBlank = allow_blank
        self.errorStyle, self.errorTitle, self.error = errorStyle, errorTitle, error
        self.showErrorMessage = showErrorMessage
        self.sqref = []

    def add(self, ref):
        self.sqref.append(str(ref))


class _写格:
    def __init__(self):
        self.value = None
        self.font = None
        self.fill = None
        self.number_format = None
        self.hyperlink = None
        self.alignment = None


class _写维:
    def __init__(self):
        self.height = None
        self.width = None
        self.hidden = False
        self.outlineLevel = 0
        self.collapsed = False


class _写条件集:
    def __init__(self, 表):
        self._表 = 表

    def add(self, sqref, rule):
        self._表._条件.append((str(sqref), rule))

    def __bool__(self):
        return bool(self._表._条件)


class _写表:
    def __init__(self, title):
        self.title = title
        self._格 = {}
        self._最大行 = 0
        self._最大列 = 0
        self.row_dimensions = defaultdict(_写维)
        self.column_dimensions = defaultdict(_写维)
        self.freeze_panes = None
        self._验证 = []
        self._条件 = []
        self._合并 = set()
        self.conditional_formatting = _写条件集(self)

    def merge_cells(self, 区域):
        self._合并.add(str(区域))

    def cell(self, row, column, value=None):
        g = self._格.setdefault((row, column), _写格())
        if value is not None:
            g.value = value
        self._最大行 = max(self._最大行, row)
        self._最大列 = max(self._最大列, column)
        return g

    def append(self, 行值):
        r = self._最大行 + 1
        self._最大行 = r
        for c, v in enumerate(行值, 1):
            if v is not None:
                self.cell(r, c).value = v

    def __setitem__(self, 坐标, 值):
        行, 列 = 拆坐标(坐标)
        self.cell(行, 列).value = 值

    def __getitem__(self, 坐标):
        行, 列 = 拆坐标(坐标)
        return self.cell(行, 列)

    def add_data_validation(self, dv):
        self._验证.append(dv)


def _格样式自写格(g):
    样 = 格样式()
    有 = False
    if g.font is not None:
        样.字体名, 样.字号 = g.font.name, g.font.size
        样.加粗 = bool(g.font.bold)
        if g.font.color:
            样.字色 = g.font.color if isinstance(g.font.color, str) else getattr(g.font.color, "rgb", None)
        有 = True
    if g.fill is not None:
        样.填充类型 = g.fill.patternType
        色 = g.fill.fgColor
        样.填充色 = 色 if isinstance(色, str) else getattr(色, "rgb", None)
        有 = True
    if g.number_format:
        样.数字格式 = g.number_format
        有 = True
    if g.hyperlink:
        样.超链接 = g.hyperlink if isinstance(g.hyperlink, str) else getattr(g.hyperlink, "target", None)
        有 = True
    if g.alignment is not None:
        样.水平, 样.垂直, 样.自动换行 = g.alignment.horizontal, g.alignment.vertical, g.alignment.wrap_text
        有 = True
    return 样 if 有 else None


def _转计划(ws):
    最大列 = max(ws._最大列, 1)
    # 行维度（行高/分级等）可能落在无格子的行上，补 None 行让写簿把行属性落进去
    底行 = max([ws._最大行, *ws.row_dimensions.keys()], default=1)
    列名 = []
    for c in range(1, 最大列 + 1):
        g = ws._格.get((1, c))
        列名.append(g.value if g is not None and g.value is not None else "")
    数据 = [[ws._格[(r, c)].value if (r, c) in ws._格 else None
             for c in range(1, 最大列 + 1)]
            for r in range(2, 底行 + 1)]
    df = pd.DataFrame(数据, columns=列名, dtype=object)
    tp = 表计划(名称=ws.title, 表=df)
    tp.冻结 = ws.freeze_panes
    tp.合并 = set(ws._合并)
    for r, d in ws.row_dimensions.items():
        rp = tp.行计划取(r)
        rp.高, rp.隐藏, rp.outline, rp.收起 = d.height, d.hidden, d.outlineLevel, d.collapsed
    for 字母, d in ws.column_dimensions.items():
        cp = tp.列计划取(字母转列(str(字母)) - 1)
        cp.宽, cp.隐藏 = d.width, d.hidden
    for (r, c), g in ws._格.items():
        样 = _格样式自写格(g)
        if 样 is not None:
            tp.盖格(r, c, 样)
    for dv in ws._验证:
        tp.验证.append({"type": dv.type, "formula1": dv.formula1,
                      "allow_blank": dv.allowBlank, "sqref": " ".join(dv.sqref),
                      "errorStyle": dv.errorStyle, "errorTitle": dv.errorTitle,
                      "error": dv.error, "showErrorMessage": dv.showErrorMessage})
    for sqref, rule in ws._条件:
        填充 = rule.fill.fgColor if rule.fill is not None else None
        if 填充 is not None and not isinstance(填充, str):
            填充 = getattr(填充, "rgb", None)
        tp.条件格式.append((sqref, 条件规则(公式=rule.formula, 填充色=填充)))
    return tp


class Workbook:
    def __init__(self):
        self._表 = [_写表("Sheet")]
        self._活动 = 0

    @property
    def active(self):
        return self._表[self._活动]

    @active.setter
    def active(self, 值):
        self._活动 = 值 if isinstance(值, int) else self._表.index(值)

    def create_sheet(self, title=None):
        t = _写表(title or f"Sheet{len(self._表) + 1}")
        self._表.append(t)
        return t

    def __getitem__(self, 名):
        for t in self._表:
            if t.title == 名:
                return t
        raise KeyError(名)

    def save(self, path):
        计划 = 簿计划()
        for ws in self._表:
            计划.表.append(_转计划(ws))
        计划.active = self._活动
        写簿(计划, str(path))
