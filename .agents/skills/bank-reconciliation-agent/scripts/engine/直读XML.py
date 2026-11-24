# -*- coding: utf-8 -*-
"""直读XML：纯标准库读取 xlsx，与 直写XML.py 镜像对称。

不经 openpyxl/pandas：zipfile 解包 + ElementTree 解析，
工作表 XML 用 iterparse 流式逐行，几十万行不整树驻留内存。

读法分两档：
- 读簿(path)：值 + 公式原文（公式格返回 "=SUM(...)" 形式）
- 读簿(path, 数据=True)：公式格改取缓存计算结果（Excel 保存后才有）

样式按"格子自身 s → 列样式兜底 → 0 号默认"还原有效格式，
与 Excel 的显示规则一致（列样式兜底是直写端快速通道的关键写法）。
"""
from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

_HEADER_RE = re.compile(r"^([A-Z]+)([0-9]+)$")
_公式片段 = re.compile(
    r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|\[[^\]]*\]|'
    r'(?<![\w.])(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*|'
    r'\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?![\w.(!])')


def _平移共享公式(公式, 起点, 终点):
    """按 A1 引用规则展开 Excel 保存的共享公式，保留引号、表名和绝对引用。"""
    起行, 起列 = 拆坐标(起点)
    终行, 终列 = 拆坐标(终点)
    def 平移地址(地址):
        m = re.fullmatch(r'(\$?)([A-Za-z]+)?(\$?)(\d+)?', 地址)
        列锁, 列文, 行锁, 行文 = m.groups()
        if 列文:
            列 = 字母转列(列文)
            if 列 > 16384:
                return 地址
            列 += 0 if 列锁 else 终列 - 起列
            if not 1 <= 列 <= 16384:
                return '#REF!'
            列文 = 列转字母(列)
        if 行文:
            行 = int(行文) + (0 if (行锁 or (not 列文 and 列锁)) else 终行 - 起行)
            if not 1 <= 行 <= 1048576:
                return '#REF!'
            行文 = str(行)
        return 列锁 + (列文 or '') + 行锁 + (行文 or '')
    def 替换(m):
        文 = m.group()
        if 文[0] in ('"', "'", '['):
            return 文
        return ':'.join(平移地址(x) for x in 文.split(':'))
    return _公式片段.sub(替换, 公式)

# openpyxl 认定的内置日期格式编号
_内置日期格式 = frozenset(range(14, 23)) | frozenset(range(27, 37)) | frozenset(range(45, 48)) | frozenset(range(50, 59))
_内置格式 = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
    9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
    14: "yyyy-mm-dd", 49: "@",
}
_日期记号 = re.compile(r"(?<![a-z])[ymdh秒分日月年]+(?![a-z])", re.IGNORECASE)
_日期排除 = re.compile(r"[\[\]\\*_]|0\.|#|,|General", re.IGNORECASE)


def 字母转列(字母: str) -> int:
    列 = 0
    for ch in 字母.upper():
        列 = 列 * 26 + (ord(ch) - 64)
    return 列


def 列转字母(列: int) -> str:
    片 = ""
    while 列:
        列, 余 = divmod(列 - 1, 26)
        片 = chr(65 + 余) + 片
    return 片


def 拆坐标(ref: str) -> tuple[int, int]:
    """A1 → (行号, 列号)，均 1 基。"""
    m = _HEADER_RE.match(ref)
    if not m:
        raise ValueError(f"坐标写法不对: {ref}")
    return int(m.group(2)), 字母转列(m.group(1))


def _像日期格式(代码: str) -> bool:
    if not 代码 or _日期排除.search(代码):
        return False
    return bool(_日期记号.search(代码))


def _序号转日期(值: float):
    基 = datetime(1899, 12, 30)
    return 基 + timedelta(days=值)


class _样式库:
    """styles.xml 还原：字体/填充/边框/数字格式/对齐索引表。"""

    def __init__(self, z: zipfile.ZipFile):
        self.数字格式 = dict(_内置格式)
        self.字体 = []           # (名, 号, 粗, 色, 下划线)
        self.填充 = []           # (类型, 色)
        self.边框 = []           # (上, 下, 左, 右)，各为 (样式, 色) 或 None
        self.xfs = []            # (numFmtId, fontId, fillId, borderId, 对齐)
        self.dxfs = []           # (字色, 填充色)
        self.日期xf = set()
        try:
            root = ET.fromstring(z.read("xl/styles.xml"))
        except KeyError:
            return
        节点 = root.find(_NS + "numFmts")
        if 节点 is not None:
            for n in 节点.findall(_NS + "numFmt"):
                self.数字格式[int(n.get("numFmtId"))] = n.get("formatCode")
        节点 = root.find(_NS + "fonts")
        for f in (节点.findall(_NS + "font") if 节点 is not None else []):
            名 = f.find(_NS + "name")
            号 = f.find(_NS + "sz")
            色 = f.find(_NS + "color")
            色值 = None
            if 色 is not None:
                色值 = 色.get("rgb") or (f"theme:{色.get('theme')}" if 色.get("theme") is not None else None)
            self.字体.append((
                名.get("val") if 名 is not None else None,
                float(号.get("val")) if 号 is not None and 号.get("val") else None,
                f.find(_NS + "b") is not None,
                色值,
                (f.find(_NS + "u") is not None and (f.find(_NS + "u").get("val") or "single")) or None,
            ))
        节点 = root.find(_NS + "fills")
        for f in (节点.findall(_NS + "fill") if 节点 is not None else []):
            pf = f.find(_NS + "patternFill")
            if pf is None:
                self.填充.append((None, None))
                continue
            fg = pf.find(_NS + "fgColor")
            self.填充.append((pf.get("patternType"), fg.get("rgb") if fg is not None else None))
        节点 = root.find(_NS + "borders")
        for b in (节点.findall(_NS + "border") if 节点 is not None else []):
            边 = {}
            for 位 in ("top", "bottom", "left", "right"):
                el = b.find(_NS + 位)
                if el is not None and el.get("style"):
                    色 = el.find(_NS + "color")
                    边[位] = (el.get("style"), 色.get("rgb") if 色 is not None else None)
            self.边框.append((边.get("top"), 边.get("bottom"), 边.get("left"), 边.get("right")))
        节点 = root.find(_NS + "cellXfs")
        for xf in (节点.findall(_NS + "xf") if 节点 is not None else []):
            al = xf.find(_NS + "alignment")
            对齐 = (al.get("horizontal"), al.get("vertical"), al.get("wrapText") == "1") if al is not None else (None, None, False)
            fmt_id = int(xf.get("numFmtId", 0))
            self.xfs.append((
                fmt_id,
                int(xf.get("fontId", 0)),
                int(xf.get("fillId", 0)),
                int(xf.get("borderId", 0)),
                对齐,
            ))
            if fmt_id in _内置日期格式 or _像日期格式(self.数字格式.get(fmt_id, "")):
                self.日期xf.add(len(self.xfs) - 1)
        节点 = root.find(_NS + "dxfs")
        for d in (节点.findall(_NS + "dxf") if 节点 is not None else []):
            色 = d.find(f"{_NS}font/{_NS}color")
            填 = d.find(f"{_NS}fill/{_NS}fgColor")
            self.dxfs.append((色.get("rgb") if 色 is not None else None,
                              填.get("rgb") if 填 is not None else None))

    def 样式(self, 索引: int) -> dict:
        """把 xf 索引还原成可读的有效格式描述。"""
        if not self.xfs or not (0 <= 索引 < len(self.xfs)):
            return {"字体": (None, None, False, None, None), "填充": (None, None),
                    "边框": (None, None, None, None), "对齐": (None, None, False),
                    "数字格式": "General"}
        fmt_id, font_id, fill_id, border_id, 对齐 = self.xfs[索引]
        return {
            "字体": self.字体[font_id] if font_id < len(self.字体) else None,
            "填充": self.填充[fill_id] if fill_id < len(self.填充) else None,
            "边框": self.边框[border_id] if border_id < len(self.边框) else None,
            "对齐": 对齐,
            "数字格式": self.数字格式.get(fmt_id, fmt_id),
        }

    def 是日期(self, 索引: int) -> bool:
        return 索引 in self.日期xf


class 表读:
    """单张工作表的读取结果。"""

    def __init__(self, 名称, 状态):
        self.名称 = 名称
        self.状态 = 状态
        self.tab色 = None
        self.冻结 = None              # 锚点（"A2" 形式）或 None
        self.筛选 = None
        self.网格线 = True
        self.缩放 = None
        self.左上角 = None
        self.outline汇总在下 = None
        self.显示outline符号 = None
        self.打印 = {}
        self.合并 = set()
        self.列 = {}                  # 列号1基 → {"宽", "隐藏", "样式"}
        self.行 = {}                  # 行号 → {"高", "隐藏", "分级", "收起"}
        self.默认行高 = None
        self.默认列宽 = None
        self.格 = {}                  # (行,列) → 值
        self.格样式索引 = {}          # (行,列) → xf 索引（仅显式写了 s 的格）
        self.公式 = {}                # (行,列) → 公式原文（不含 =）
        self.批注 = {}                # (行,列) → (文本, 作者)
        self.超链接 = {}              # (行,列) → 目标
        self.数据验证 = []            # (sqref, type, allowBlank, formula1, formula2, errorStyle, errorTitle, error, showErrorMessage)
        self.条件格式 = []            # (sqref, type, 公式们, dxf描述)
        self.最大行 = 0
        self.最大列 = 0

    def 值(self, 行: int, 列: int):
        return self.格.get((行, 列))

    def 样式(self, 行: int, 列: int, 样式库: "_样式库") -> dict:
        """有效格式：格样式 → 列样式兜底 → 0 号默认。"""
        索引 = self.格样式索引.get((行, 列))
        if 索引 is None:
            索引 = self.列.get(列, {}).get("样式", 0)
        return 样式库.样式(索引 or 0)

    def 行值(self, 起: int = 1, 止: int | None = None):
        """逐行产出 [v1, v2, ...]，长度对齐本表最大列；空格为 None。"""
        止 = 止 if 止 is not None else self.最大行
        for r in range(起, 止 + 1):
            yield [self.格.get((r, c)) for c in range(1, self.最大列 + 1)]


def _表目标(z: zipfile.ZipFile) -> list[tuple[str, str, str]]:
    """[(表名, sheet部件路径, 状态)]，按工作簿顺序。"""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    目标 = {r.get("Id"): r.get("Target") for r in rels}
    出 = []
    sheets = wb.find(_NS + "sheets")
    for i, s in enumerate(sheets.findall(_NS + "sheet") if sheets is not None else [], 1):
        t = (目标.get(s.get(_REL + "id")) or f"worksheets/sheet{i}.xml").lstrip("/")
        出.append((s.get("name"), t if t.startswith("xl/") else "xl/" + t, s.get("state", "visible")))
    return 出


def _读共享串(z: zipfile.ZipFile) -> list[str]:
    try:
        流 = z.open("xl/sharedStrings.xml")
    except KeyError:
        return []
    串 = []
    for _事件, 元 in ET.iterparse(流, events=("end",)):
        if 元.tag == _NS + "si":
            串.append("".join(t.text or "" for t in 元.iter(_NS + "t")))
            元.clear()
    流.close()
    return 串


def _颜色(属性: dict) -> str | None:
    if "rgb" in 属性:
        return 属性["rgb"]
    if "theme" in 属性:
        return f"theme:{属性['theme']}"
    if "indexed" in 属性:
        return f"indexed:{属性['indexed']}"
    return None


class 簿读:
    """整簿读取结果：簿读['表名'] → 表读。"""

    def __init__(self, path, *, 数据: bool = False, 读样式: bool = True):
        self.路径 = str(path)
        self.数据模式 = 数据
        self.表 = {}
        self.表名列表 = []
        self.计算设置 = {}
        self.active名 = None
        self.样式库 = None
        with zipfile.ZipFile(self.路径) as z:
            self.样式库 = _样式库(z) if 读样式 else None
            sst = _读共享串(z)
            try:
                wb = ET.fromstring(z.read("xl/workbook.xml"))
                calc = wb.find(_NS + "calcPr")
                if calc is not None:
                    self.计算设置 = dict(calc.attrib)
                view = wb.find(f"{_NS}bookViews/{_NS}workbookView")
                活动 = int(view.get("activeTab", 0)) if view is not None else 0
                定义 = wb.find(_NS + "definedNames")
                自定 = {}            # sheet序(0基) → {名: 值}
                if 定义 is not None:
                    for d in 定义.findall(_NS + "definedName"):
                        局部 = d.get("localSheetId")
                        if 局部 is None:
                            continue
                        自定.setdefault(int(局部), {})[d.get("name") or ""] = d.text or ""
            except KeyError:
                活动 = 0
                自定 = {}
            for 序, (名, 目标, 状态) in enumerate(_表目标(z)):
                t = self._读表(z, 目标, sst)
                t.名称, t.状态 = 名, 状态
                for 名2, 值 in 自定.get(序, {}).items():
                    # 打印定义名：'表名'!$1:$1 / '表名'!$A$1:$M$78 → 1:1 / A1:M78
                    体 = 值.split("!")[-1].replace("$", "")
                    if 名2.endswith("Print_Titles"):
                        t.打印["title_rows"] = 体
                    elif 名2.endswith("Print_Area"):
                        t.打印["area"] = 体
                self._读表附属(z, 目标, t)
                self.表[名] = t
                self.表名列表.append(名)
                if 序 == 活动:
                    self.active名 = 名

    def __getitem__(self, 名: str) -> 表读:
        return self.表[名]

    def __contains__(self, 名: str) -> bool:
        return 名 in self.表

    def __iter__(self):
        return iter(self.表.values())

    # ---- 内部 ----

    def _单元值(self, c, sst):
        f = c.find(_NS + "f")
        if f is not None and not self.数据模式:
            return "=" + (f.text or "")
        v = c.find(_NS + "v")
        if v is None or v.text is None:
            if f is not None:
                return None        # 数据模式但无缓存值
            is_ = c.find(_NS + "is")
            if is_ is not None:    # inlineStr
                return "".join(t.text or "" for t in is_.iter(_NS + "t"))
            return None
        t = c.get("t")
        if t == "s":
            return sst[int(v.text)]
        if t == "b":
            return v.text == "1"
        if t in ("str", "e"):
            return v.text
        try:
            数 = float(v.text)
        except (TypeError, ValueError):
            return v.text
        s = c.get("s")
        if self.样式库 is not None and s is not None and self.样式库.是日期(int(s)):
            return _序号转日期(数)
        return int(数) if 数 == int(数) and "e" not in v.text.lower() and "." not in v.text else 数

    def _读表(self, z, 目标, sst) -> 表读:
        t = 表读("", "visible")
        流 = z.open(目标)
        状 = None
        共享定义, 共享引用 = {}, []
        for 事件, 元 in ET.iterparse(流, events=("start", "end")):
            tag = 元.tag
            if 事件 == "start":
                if tag == _NS + "sheetPr":
                    状 = "sheetPr"
                elif tag == _NS + "pageSetUpPr" and 状 == "sheetPr":
                    t.打印["fitToPage"] = 元.get("fitToPage") in ("1", "true")
                elif tag == _NS + "tabColor" and 状 == "sheetPr":
                    t.tab色 = _颜色(元.attrib)
                elif tag == _NS + "outlinePr":
                    t.outline汇总在下 = 元.get("summaryBelow") != "0"
                elif tag == _NS + "sheetView":
                    t.网格线 = 元.get("showGridLines") != "0"
                    if 元.get("zoomScale"):
                        t.缩放 = int(元.get("zoomScale"))
                    if 元.get("topLeftCell"):
                        t.左上角 = 元.get("topLeftCell")
                elif tag == _NS + "pane":
                    锚 = 元.get("topLeftCell")
                    if 锚 and 元.get("state") in ("frozen", "frozenSplit"):
                        t.冻结 = 锚
                elif tag == _NS + "sheetFormatPr":
                    if 元.get("defaultRowHeight"):
                        t.默认行高 = float(元.get("defaultRowHeight"))
                    if 元.get("defaultColWidth"):
                        t.默认列宽 = float(元.get("defaultColWidth"))
                elif tag == _NS + "col":
                    起, 讫 = int(元.get("min")), int(元.get("max"))
                    for i in range(起, 讫 + 1):
                        t.列[i] = {
                            "宽": float(元.get("width")) if 元.get("width") else None,
                            "隐藏": 元.get("hidden") == "1",
                            "样式": int(元.get("style", 0)),
                        }
                        t.最大列 = max(t.最大列, i)
                elif tag == _NS + "row":
                    状 = int(元.get("r", 0))
                    行属 = {}
                    if 元.get("ht"):
                        行属["高"] = float(元.get("ht"))
                    if 元.get("hidden") == "1":
                        行属["隐藏"] = True
                    if 元.get("outlineLevel"):
                        行属["分级"] = int(元.get("outlineLevel"))
                    if 元.get("collapsed") == "1":
                        行属["收起"] = True
                    if 行属:
                        t.行[状] = {"高": None, "隐藏": False, "分级": 0, "收起": False, **行属}
                    t.最大行 = max(t.最大行, 状)
                elif tag == _NS + "mergeCell":
                    t.合并.add(元.get("ref"))
                elif tag == _NS + "autoFilter":
                    t.筛选 = 元.get("ref")
                elif tag == _NS + "dataValidation":
                    f1 = 元.find(f"{_NS}formula1")
                    f2 = 元.find(f"{_NS}formula2")
                    t.数据验证.append((
                        元.get("sqref"), 元.get("type"), 元.get("allowBlank") == "1",
                        f1.text if f1 is not None else None,
                        f2.text if f2 is not None else None,
                        元.get("errorStyle"), 元.get("errorTitle"), 元.get("error"),
                        元.get("showErrorMessage") == "1",
                    ))
                elif tag == _NS + "conditionalFormatting":
                    状 = ("cf", 元.get("sqref"))
                elif tag == _NS + "pageMargins":
                    t.打印["margins"] = dict(元.attrib)
                elif tag == _NS + "pageSetup":
                    t.打印["setup"] = dict(元.attrib)
                elif tag == _NS + "hyperlink":
                    ref = 元.get("ref")
                    位置 = 元.get('location')
                    目标2 = ('#' + 位置) if 位置 is not None else (元.get(_REL + "id") or "")
                    try:
                        首, _, 尾 = ref.partition(':')
                        首行, 首列 = 拆坐标(首)
                        尾行, 尾列 = 拆坐标(尾 or 首)
                        for 行 in range(首行, 尾行 + 1):
                            for 列 in range(首列, 尾列 + 1):
                                t.超链接[(行, 列)] = 目标2
                    except ValueError:
                        pass
            else:
                if tag == _NS + "c":
                    ref = 元.get("r")
                    try:
                        行列 = 拆坐标(ref)
                    except ValueError:
                        行列 = None
                    if 行列 is not None:
                        值 = self._单元值(元, sst)
                        if 值 is not None:
                            t.格[行列] = 值
                        f = 元.find(_NS + "f")
                        if f is not None:
                            t.公式[行列] = f.text or ""
                            if f.get('t') == 'shared':
                                if f.text:
                                    共享定义[f.get('si')] = (ref, f.text)
                                else:
                                    共享引用.append((行列, ref, f.get('si')))
                        if 元.get("s"):
                            t.格样式索引[行列] = int(元.get("s"))
                        t.最大行 = max(t.最大行, 行列[0])
                        t.最大列 = max(t.最大列, 行列[1])
                    元.clear()
                elif tag == _NS + "row":
                    状 = None
                    元.clear()
                elif tag == _NS + "cfRule" and isinstance(状, tuple):
                    公式们 = tuple((x.text or "") for x in 元.findall(_NS + "formula"))
                    dxfid = int(元.get("dxfId", -1))
                    dxf描述 = (self.样式库.dxfs[dxfid]
                               if self.样式库 is not None and 0 <= dxfid < len(self.样式库.dxfs)
                               else None)
                    t.条件格式.append((状[1], 元.get("type"), 公式们, dxf描述))
                    元.clear()
                elif tag == _NS + "conditionalFormatting":
                    状 = None
                elif tag == _NS + "sheetPr":
                    状 = None
        流.close()
        for 行列, ref, si in 共享引用:
            if si not in 共享定义:
                raise ValueError(f'{目标}!{ref} 缺少共享公式定义 {si}')
            起点, 公式 = 共享定义[si]
            展开 = _平移共享公式(公式, 起点, ref)
            t.公式[行列] = 展开
            if not self.数据模式:
                t.格[行列] = '=' + 展开
        return t

    def _读表附属(self, z, 目标, t: 表读):
        """sheet rels：批注、超链接目标还原。"""
        rel路径 = 目标.replace("worksheets/", "worksheets/_rels/") + ".rels"
        try:
            rels = ET.fromstring(z.read(rel路径))
        except KeyError:
            return
        rid目标 = {}
        批注目标 = None
        for r in rels:
            型 = r.get("Type") or ""
            if 型.endswith("/hyperlink"):
                rid目标[r.get("Id")] = r.get("Target")
            elif 型.endswith("/comments"):
                批注目标 = r.get("Target")
        for 行列, 目标2 in list(t.超链接.items()):
            if 目标2 in rid目标:
                t.超链接[行列] = rid目标[目标2]
        if 批注目标:
            批注路径 = "xl/" + 批注目标.replace("../", "").lstrip("/")
            try:
                root = ET.fromstring(z.read(批注路径))
            except KeyError:
                return
            作者 = [a.text or "" for a in root.iter(_NS + "author")]
            for c in root.iter(_NS + "comment"):
                try:
                    行列 = 拆坐标(c.get("ref"))
                except ValueError:
                    continue
                文 = "".join(x.text or "" for x in c.iter(_NS + "t"))
                aid = int(c.get("authorId", 0))
                t.批注[行列] = (文, 作者[aid] if aid < len(作者) else "")


def 读表名(path) -> list[str]:
    """只读 workbook.xml 取表名列表，不解析各表内容。"""
    with zipfile.ZipFile(str(path)) as z:
        return [名 for 名, _目标, _状态 in _表目标(z)]


def 读簿(path, *, 数据: bool = False, 读样式: bool = True) -> 簿读:
    """读取 xlsx：默认公式格返回公式原文；数据=True 时取缓存计算结果。"""
    return 簿读(path, 数据=数据, 读样式=读样式)


def 读表数据(path, 表=0, *, 跳过行: int = 0, 限行: int | None = None):
    """轻量入口：读一张表的值网格（公式格取缓存值），供数据装载。

    表：0 基序号或表名。返回 (列数, 行迭代器)，每行是长度=列数的列表。
    空格为 None；日期格式数字格已转为 datetime。
    """
    簿 = 读簿(path, 数据=True)
    t = 簿.表[簿.表名列表[表]] if isinstance(表, int) else 簿.表[表]
    止 = t.最大行 if 限行 is None else min(t.最大行, 跳过行 + 限行)
    return t.最大列, t.行值(起=跳过行 + 1, 止=止)
