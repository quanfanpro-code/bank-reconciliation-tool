# -*- coding: utf-8 -*-
"""直写XML：把 排版计划.簿计划 直接序列化为 xlsx（OpenXML）。

不经 openpyxl 对象模型：样式集中到 styles.xml 索引表（cellXfs），文本去重进
sharedStrings，公式格 <f> 不带缓存值、calcPr fullCalcOnLoad 打开时强制重算。
结构规范照 minimax-xlsx 技能的 minimal_xlsx 模板（D:/BaiduSyncdisk/workbuddy skills）。
"""
import math
import zipfile
from datetime import datetime as _dt, date as _date
from xml.sax.saxutils import escape as _xesc

from 排版计划 import 合并样式, 簿计划

_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

_列字母 = {}


def _字母(col):
    """1 基列号 → Excel 字母（带缓存）。"""
    s = _列字母.get(col)
    if s is None:
        n = col
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        _列字母[col] = s
    return s


def _rgb(色):
    """RRGGBB → 00RRGGBB（minimax 模板约定：00 alpha = 不透明）。"""
    return 色 if len(色) == 8 else "00" + 色


def _串文本(v):
    return v if isinstance(v, str) else str(v)


class _样式表:
    """styles.xml 构建：字体/填充/边框/数字格式/xf 全部按内容去重。"""

    def __init__(self):
        self._numfmt = {}          # 代码 → id（0=General、49=@、10=0.00% 用内置）
        self._numfmt列表 = []
        self._font = {}            # 键 → id
        self._font列表 = []
        self._fill = {}
        self._fill列表 = []
        self._border = {}
        self._border列表 = []
        self._xf = {}              # xf 键 → 索引（0 为默认，预置）
        self._xf列表 = []
        self._dxf = {}             # 条件格式 dxf 键 → id
        self._dxf列表 = []
        # 规范约定（与 openpyxl 产物一致）：fonts[0] 是默认字体 Calibri 11 +
        # 主题色 1；"未显式指定字体"的格都落到它上面，读回才是 Calibri 11
        self.font("Calibri", 11, None, "theme:1", None)
        # fills[0]=none、fills[1]=gray125 同理占位
        self.fill(None, None)
        self._fill[("gray125", None)] = 1
        self._fill列表.append('<fill><patternFill patternType="gray125"/></fill>')

    # ---- 数字格式 ----
    def numfmt(self, 代码):
        if 代码 in (None, "", "General"):
            return 0
        if 代码 == "@":
            return 49
        if 代码 == "0.00%":
            return 10
        if 代码 == "0.00":
            return 2
        i = self._numfmt.get(代码)
        if i is None:
            i = 164 + len(self._numfmt列表)
            self._numfmt[代码] = i
            self._numfmt列表.append(代码)
        return i

    def font(self, 名, 号, 粗, 色, 下划线):
        if 名 is None and 号 is None and 粗 is None and 色 is None and 下划线 is None:
            return 0                    # 0 号位 = 默认字体（Calibri 11 主题色 1）
        键 = (名, 号, 粗, 色, 下划线)
        i = self._font.get(键)
        if i is None:
            片 = ["<font>"]
            if 粗:
                片.append("<b/>")
            if 号 is not None:
                片.append(f'<sz val="{号:g}"/>')
            if 名:
                片.append(f'<name val="{_xesc(名)}"/>')
            if 色 == "theme:1":
                片.append('<color theme="1"/>')
            elif 色:
                片.append(f'<color rgb="{_rgb(色)}"/>')
            if 下划线:
                片.append(f'<u val="{_xesc(下划线)}"/>')
            片.append("</font>")
            i = len(self._font列表)
            self._font[键] = i
            self._font列表.append("".join(片))
        return i

    def fill(self, 类型, 色):
        if 类型 in (None, "none"):
            键, 片 = ("none", None), '<fill><patternFill patternType="none"/></fill>'
        elif 类型 == "gray125" and not 色:
            键, 片 = ("gray125", None), '<fill><patternFill patternType="gray125"/></fill>'
        else:
            键 = (类型, 色)
            片 = (f'<fill><patternFill patternType="{_xesc(类型 or "solid")}">'
                  f'<fgColor rgb="{_rgb(色 or "00000000")}"/><bgColor indexed="64"/>'
                  "</patternFill></fill>")
        i = self._fill.get(键)
        if i is None:
            i = len(self._fill列表)
            self._fill[键] = i
            self._fill列表.append(片)
        return i

    def border(self, 上, 下, 左, 右):
        键 = (上, 下, 左, 右)
        i = self._border.get(键)
        if i is None:
            片 = ["<border>"]
            for 边名, 框 in (("left", 左), ("right", 右), ("top", 上), ("bottom", 下)):
                if not 框 or (isinstance(框, tuple) and not 框[0]):
                    片.append(f"<{边名}/>")
                else:
                    式, 色 = 框
                    c = f'<color rgb="{_rgb(色)}"/>' if 色 else ""
                    片.append(f'<{边名} style="{_xesc(式)}">{c}</{边名}>')
            片.append("<diagonal/></border>")
            i = len(self._border列表)
            self._border[键] = i
            self._border列表.append("".join(片))
        return i

    def xf(self, s: "格样式 | None"):
        """格样式 → cellXfs 索引；批注/超链接不参与样式键（另行处理）。"""
        from 排版计划 import 格样式
        s = s or 格样式()
        fmt = self.numfmt(s.数字格式)
        f = self.font(s.字体名 or None, s.字号, bool(s.加粗) if s.加粗 is not None else None,
                      s.字色, s.下划线)
        fl = self.fill(s.填充类型, s.填充色)
        b = self.border(s.上框, s.下框, s.左框, s.右框)
        对齐 = (s.水平 or None, s.垂直 or None,
                bool(s.自动换行) if s.自动换行 is not None else None)
        键 = (fmt, f, fl, b, 对齐)
        i = self._xf.get(键)
        if i is None:
            片 = [f'<xf numFmtId="{fmt}" fontId="{f}" fillId="{fl}" borderId="{b}" xfId="0"']
            片.append(' applyNumberFormat="1"' if fmt else "")
            片.append(' applyFont="1"' if f else "")
            片.append(' applyFill="1"' if fl else "")
            片.append(' applyBorder="1"' if b else "")
            if any(对齐):
                片.append(' applyAlignment="1"><alignment')
                if 对齐[0]:
                    片.append(f' horizontal="{_xesc(对齐[0])}"')
                if 对齐[1]:
                    片.append(f' vertical="{_xesc(对齐[1])}"')
                if 对齐[2]:
                    片.append(' wrapText="1"')
                片.append("/></xf>")
            else:
                片.append("/>")
            i = len(self._xf列表)
            self._xf[键] = i
            self._xf列表.append("".join(片))
        return i

    def dxf(self, 规则):
        """自有条件规则 → dxf 索引。"""
        片 = ["<dxf>"]
        if 规则.字色:
            片.append(f'<font><color rgb="{_rgb(规则.字色)}"/></font>')
        if 规则.填充色:
            # dxf 里的纯色填充：openpyxl 旧产物把颜色放 bgColor（Excel 的 dxf 约定）
            片.append("<fill><patternFill>"
                      f'<bgColor rgb="{_rgb(规则.填充色)}"/>'
                      "</patternFill></fill>")
        片.append("</dxf>")
        键 = (规则.字色, 规则.填充色)
        i = self._dxf.get(键)
        if i is None:
            i = len(self._dxf列表)
            self._dxf[键] = i
            self._dxf列表.append("".join(片))
        return i

    def xml(self):
        片 = [_HEADER, f"<styleSheet {_NS}>"]
        if self._numfmt列表:
            片.append(f'<numFmts count="{len(self._numfmt列表)}">')
            for 代码 in self._numfmt列表:
                i = self._numfmt[代码]
                片.append(f'<numFmt numFmtId="{i}" formatCode="{_xesc(代码)}"/>')
            片.append("</numFmts>")
        片.append(f'<fonts count="{max(1, len(self._font列表))}">')
        片.extend(self._font列表 or ["<font><sz val='11'/><name val='Calibri'/></font>"])
        片.append("</fonts>")
        片.append(f'<fills count="{len(self._fill列表)}">')
        片.extend(self._fill列表)
        片.append("</fills>")
        片.append(f'<borders count="{max(1, len(self._border列表))}">')
        片.extend(self._border列表 or ["<border><left/><right/><top/><bottom/><diagonal/></border>"])
        片.append("</borders>")
        片.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        片.append(f'<cellXfs count="{len(self._xf列表)}">')
        片.extend(self._xf列表 or ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'])
        片.append("</cellXfs>")
        片.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
        if self._dxf列表:
            片.append(f'<dxfs count="{len(self._dxf列表)}">')
            片.extend(self._dxf列表)
            片.append("</dxfs>")
        片.append("</styleSheet>")
        return "".join(片)


def _数字文本(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        v = round(v, 15)                  # 与 openpyxl 写盘一致：抹掉浮点尾噪声
        return repr(v) if v != int(v) or abs(v) >= 1e15 else str(int(v))
    if isinstance(v, _dt):
        基 = _dt(1899, 12, 30)
        秒 = (v - 基).total_seconds()
        return repr(秒 / 86400.0)
    if isinstance(v, _date):
        return str((v - _date(1899, 12, 30)).days)
    return None


def _冻结解析(冻结):
    if not 冻结:
        return None
    数字 = "".join(ch for ch in 冻结 if ch.isdigit())
    字母 = "".join(ch for ch in 冻结 if ch.isalpha())
    行 = int(数字 or 0)
    列 = 0
    for ch in 字母:
        列 = 列 * 26 + (ord(ch.upper()) - 64)
    return 列, 行


class _表写出:
    """单张表 → sheet XML 文本 + 批注表。"""

    def __init__(self, tp, 样式表, sst, sst序):
        self.tp = tp
        self.样式表 = 样式表
        self.sst = sst
        self.sst序 = sst序
        self.批注 = []             # (引用, 文本, 作者)

    def 串(self, v):
        i = self.sst.get(v)
        if i is None:
            i = len(self.sst序)
            self.sst[v] = i
            self.sst序.append(v)
        return i

    def 格(self, 行号, 列号, s, v, 强制空=False, 省样式=False):
        ref = f"{_字母(列号)}{行号}"
        if 省样式:
            属 = f' r="{ref}"'
        else:
            索引 = self.样式表.xf(s) if s is not None else 0
            属 = f' r="{ref}"' + (f' s="{索引}"' if (索引 or s is not None) else "")
        if v is None:
            if 强制空:
                return f"<c{属}/>"
            return None
        if isinstance(v, str) and v.startswith("="):
            return f'<c{属}><f>{_xesc(v[1:])}</f><v></v></c>'
        if isinstance(v, str):
            return f'<c{属} t="s"><v>{self.串(v)}</v></c>'
        if isinstance(v, bool):
            return f'<c{属} t="b"><v>{1 if v else 0}</v></c>'
        n = _数字文本(v)
        if n is None:
            return None
        return f"<c{属}><v>{n}</v></c>"

    def 出(self):
        tp = self.tp
        偏 = tp.前置空列
        df = tp.表
        from 报告写入 import _值转换  # 复用同一套空值语义
        from 排版计划 import 格样式, _快速通道行数

        # ---- 大表列默认（与报告写入同语义：列基础样式+格变换按列代表值） ----
        大表 = len(df) > _快速通道行数
        列默认 = []
        if 大表 and tp.格变换 is not None:
            for j in range(len(df.columns)):
                s = tp.列[j].基础样式 if j < len(tp.列) else None
                s = 合并样式(None, s)
                代表 = next((x for x in df.iloc[:, j]
                             if x is not None
                             and not (isinstance(x, float) and math.isnan(x))), None)
                列默认.append(tp.格变换(s, 代表, str(df.columns[j])))
        elif 大表:
            列默认 = [(tp.列[j].基础样式 if j < len(tp.列) else None)
                      for j in range(len(df.columns))]

        行样式 = {行号: rp.样式 for 行号, rp in tp.行.items() if rp.样式 is not None}
        覆盖 = tp.覆盖
        后覆盖 = tp.后覆盖
        空缺 = tp.空缺格
        物化 = tp.物化空格

        片 = [_HEADER, f"<worksheet {_NS}>"]
        批注格 = []                # (行号, 列号1基, 文本, 作者)

        # ---- sheetPr ----
        pr = []
        if tp.tab色:
            pr.append(f'<tabColor rgb="{_rgb(tp.tab色)}"/>')
        if tp.outline汇总在下 is not None:
            pr.append(f'<outlinePr summaryBelow="{1 if tp.outline汇总在下 else 0}"/>')
        打印 = tp.打印 or {}
        if 打印.get("fitToPage"):
            pr.append('<pageSetUpPr fitToPage="1"/>')
        if pr:
            片.append(f"<sheetPr>{''.join(pr)}</sheetPr>")

        # ---- sheetViews ----
        视图 = ['<sheetView workbookViewId="0"']
        if tp.左上角:
            视图.append(f' topLeftCell="{tp.左上角}"')
        if not tp.网格线:
            视图.append(' showGridLines="0"')
        if tp.缩放 is not None:
            视图.append(f' zoomScale="{tp.缩放}" zoomScaleNormal="{tp.缩放}"')
        冻结 = _冻结解析(tp.冻结)
        if 冻结:
            列, 行 = 冻结
            视图.append(f'><pane xSplit="{列}" ySplit="{行}" '
                        f'topLeftCell="{tp.冻结}" activePane="bottomRight" state="frozen"/>')
            视图.append("</sheetView>")
        else:
            视图.append("/>")
        片.append(f"<sheetViews>{''.join(视图)}</sheetViews>")

        # ---- sheetFormatPr ----
        片.append(f'<sheetFormatPr defaultRowHeight="{tp.默认行高 if tp.默认行高 is not None else 15}"/>')

        # ---- cols ----
        列片 = []
        for k, cp in enumerate(tp.前部列):
            属 = f' min="{k + 1}" max="{k + 1}"'
            if cp.宽 is not None:
                属 += f' width="{cp.宽:g}" customWidth="1"'
            if cp.隐藏:
                属 += ' hidden="1"'
            列片.append(f"<col{属}/>")
        for i, cp in enumerate(tp.列):
            属 = f' min="{偏 + i + 1}" max="{偏 + i + 1}"'
            if cp.宽 is not None:
                属 += f' width="{cp.宽:g}" customWidth="1"'
            if cp.隐藏:
                属 += ' hidden="1"'
            if 大表 and i < len(列默认) and 列默认[i] is not None:
                属 += f' style="{self.样式表.xf(列默认[i])}"'
            列片.append(f"<col{属}/>")
        if 列片:
            片.append(f"<cols>{''.join(列片)}</cols>")

        # ---- sheetData ----
        片.append("<sheetData>")
        表头含前列 = tp.表头含前列
        行1 = ["<row r=\"1\""]
        if 1 in tp.行 and tp.行[1].高 is not None:
            行1.append(f' ht="{tp.行[1].高:g}" customHeight="1"')
        行1.append(">")
        if 表头含前列:
            for k in range(偏):
                s = 合并样式(tp.表头样式, 覆盖.get((1, k + 1)))
                t = self.格(1, k + 1, s, None, 强制空=True)
                if t:
                    行1.append(t)
        for j, 名 in enumerate(df.columns):
            s = 合并样式(tp.表头样式, 覆盖.get((1, 偏 + j + 1)))
            if s is not None and s.批注:
                批注格.append((1, 偏 + j + 1, s.批注[0], s.批注[1]))
            行1.append(self.格(1, 偏 + j + 1, s, _串文本(名)))
        行1.append("</row>")
        片.append("".join(行1))

        i = -1
        for row in df.itertuples(index=False, name=None):
            i += 1
            行号 = i + 2
            rp = tp.行.get(行号)
            rs = 行样式.get(行号)
            行片 = [f"<row r=\"{行号}\""]
            if rp is not None:
                if rp.高 is not None:
                    行片.append(f' ht="{rp.高:g}" customHeight="1"')
                if rp.隐藏:
                    行片.append(' hidden="1"')
                if rp.outline:
                    行片.append(f' outlineLevel="{rp.outline}"')
                if rp.收起:
                    行片.append(' collapsed="1"')
            行片.append(">")
            for k in range(偏):
                if (行号, k + 1) in 空缺:
                    continue
                ps = tp.前部列[k].基础样式 if k < len(tp.前部列) else None
                cov = 覆盖.get((行号, k + 1))
                if cov is not None:
                    ps = 合并样式(ps, cov)
                if ps is not None:
                    t = self.格(行号, k + 1, ps, None, 强制空=True)
                    if t:
                        行片.append(t)
            j = -1
            for v in row:
                j += 1
                列号 = 偏 + j + 1
                v = _值转换(v)
                if v is None and (行号, 列号) in 空缺:
                    continue
                cov = 覆盖.get((行号, 列号))
                s2 = 后覆盖.get((行号, 列号)) if 后覆盖 else None
                列名 = str(df.columns[j]) if j < len(df.columns) else ""
                if 大表 and rs is None and cov is None and s2 is None:
                    样 = 列默认[j] if j < len(列默认) else None
                    # 普通格不写 s 属性：样式由 <col style> 兜底（快速通道的关键）
                    t = self.格(行号, 列号, None, v, 省样式=True)
                    行片.append(t) if t else None
                    continue
                样 = tp.列[j].基础样式 if j < len(tp.列) else None
                if rs is not None:
                    样 = 合并样式(样, rs)
                if cov is not None:
                    样 = 合并样式(样, cov)
                if tp.格变换 is not None and (v is not None or 物化):
                    样 = tp.格变换(样 if 样 is not None else 格样式(), v, 列名)
                if s2 is not None:
                    样 = 合并样式(样, s2)
                超链 = 样.超链接 if 样 is not None else None
                批 = 样.批注 if 样 is not None else None
                if 批:
                    批注格.append((行号, 列号, 批[0], 批[1]))
                if 超链:
                    self.超链集 = getattr(self, "超链集", [])
                    self.超链集.append((ref_超链 := f"{_字母(列号)}{行号}", 超链))
                t = self.格(行号, 列号, 样, v, 强制空=(v is None))
                if t:
                    行片.append(t)
            行片.append("</row>")
            片.append("".join(行片))
        片.append("</sheetData>")

        # ---- autoFilter ----
        if tp.筛选:
            片.append(f'<autoFilter ref="{tp.筛选}"/>')

        # ---- conditionalFormatting ----
        优先 = 1
        for 区域, 规则 in tp.条件格式:
            dxfid = self.样式表.dxf(规则)
            公式们 = "".join(f"<formula>{_xesc(str(f_))}</formula>"
                             for f_ in (规则.公式 or []))
            片.append(f'<conditionalFormatting sqref="{_xesc(str(区域))}">'
                      f'<cfRule type="expression" dxfId="{dxfid}" priority="{优先}">'
                      f"{公式们}</cfRule></conditionalFormatting>")
            优先 += 1

        # ---- dataValidations ----
        if tp.验证:
            片.append(f'<dataValidations count="{len(tp.验证)}">')
            for kw in tp.验证:
                属 = f' type="{kw["type"]}"' if kw.get("type") else ""
                if kw.get("allow_blank"):
                    属 += ' allowBlank="1"'
                if kw.get("showErrorMessage"):
                    属 += ' showErrorMessage="1"'
                if kw.get("errorStyle"):
                    属 += f' errorStyle="{kw["errorStyle"]}"'
                if kw.get("errorTitle"):
                    属 += f' errorTitle="{_xesc(kw["errorTitle"])}"'
                if kw.get("error"):
                    属 += f' error="{_xesc(kw["error"])}"'
                if kw.get("promptTitle"):
                    属 += f' promptTitle="{_xesc(kw["promptTitle"])}"'
                if kw.get("prompt"):
                    属 += f' prompt="{_xesc(kw["prompt"])}"'
                if kw.get("showInputMessage"):
                    属 += ' showInputMessage="1"'
                属 += f' sqref="{_xesc(str(kw["sqref"]))}"'
                f1 = f"<formula1>{_xesc(str(kw['formula1']))}</formula1>" if kw.get("formula1") else ""
                片.append(f"<dataValidation{属}>{f1}</dataValidation>")
            片.append("</dataValidations>")

        # ---- hyperlinks：按 openpyxl 旧路径方式走 rels（金样/测试按 target 读回） ----
        超链 = getattr(self, "超链集", None)
        if 超链:
            片.append("<hyperlinks>")
            for k, (ref, 目标) in enumerate(超链):
                片.append(f'<hyperlink ref="{ref}" r:id="rIdH{k + 1}"/>')
            片.append("</hyperlinks>")

        # ---- 打印 ----
        # openpyxl PageMargins 默认 0.75；仅当计划显式给值才覆盖
        左 = 打印.get("左边距", 0.75)
        右 = 打印.get("右边距", 0.75)
        片.append(f'<pageMargins left="{左}" right="{右}" top="0.75" bottom="0.75" '
                  'header="0.3" footer="0.3"/>')
        if 打印:
            属 = []
            if 打印.get("orientation"):
                属.append(f' orientation="{打印["orientation"]}"')
            if 打印.get("paperSize") is not None:
                属.append(f' paperSize="{打印["paperSize"]}"')
            if 打印.get("fitToWidth") is not None:
                属.append(f' fitToWidth="{打印["fitToWidth"]}"')
            if 打印.get("fitToHeight") is not None:
                属.append(f' fitToHeight="{打印["fitToHeight"]}"')
            if 属:
                片.append(f"<pageSetup{''.join(属)}/>")

        if 批注格:
            片.append('<legacyDrawing r:id="rIdLegacy"/>')
        片.append("</worksheet>")
        self.批注格 = 批注格
        return "".join(片)


def _批注XML(批注格):
    作者们 = []
    作者序 = {}
    条 = []
    for 行号, 列号, 文本, 作者 in 批注格:
        ref = f"{_字母(列号)}{行号}"
        i = 作者序.get(作者)
        if i is None:
            i = len(作者们)
            作者序[作者] = i
            作者们.append(作者)
        条.append(f'<comment ref="{ref}" authorId="{i}">'
                  f'<text><r><t xml:space="preserve">{_xesc(文本)}</t></r></text></comment>')
    xml = (_HEADER + f"<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
           f"<authors>{''.join(f'<author>{_xesc(a)}</author>' for a in 作者们)}</authors>"
           f"<commentList>{''.join(条)}</commentList></comments>")
    vml = (
        '<xml xmlns:v="urn:schemas-microsoft-com:vml"'
        ' xmlns:o="urn:schemas-microsoft-com:office:office"'
        ' xmlns:x="urn:schemas-microsoft-com:office:excel">'
        '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
        '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"'
        ' path="m,l,21600r21600,21600xe"><v:stroke joinstyle="miter"/>'
        '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
    )
    for k, (行号, 列号, 文本, 作者) in enumerate(批注格):
        ref = f"{_字母(列号)}{行号}"
        vml += (
            f'<v:shape id="_x0000_s10{25 + k}" type="#_x0000_t202"'
            " style='position:absolute;margin-left:59pt;width:144pt;height:59pt;"
            "z-index:1;visibility:hidden' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">"
            '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>'
            '<v:path o:connecttype="none"/>'
            '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left">'
            "</div></v:textbox>"
            '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
            f"<x:Anchor>{列号},15,{行号},0,{列号 + 2},31,{行号 + 4},4</x:Anchor>"
            "<x:AutoFill>False</x:AutoFill><x:Row>" + str(行号 - 1) +
            "</x:Row><x:Column>" + str(列号 - 1) + "</x:Column></x:ClientData></v:shape>"
        )
    vml += "</xml>"
    return xml, vml


def _打印名(区域):
    """B1:M78 → $B$1:$M$78（工作表名由调用方加）。"""
    import re

    def 加(p):
        m = re.match(r"([A-Z]+)(\d+)", p.strip())
        return f"${m.group(1)}${m.group(2)}" if m else p

    return ":".join(加(x) for x in 区域.split(":"))


def 写簿(计划: 簿计划, 输出路径: str) -> None:
    """渲染簿计划的直写 XML 实现：一次 zip 成型。"""
    样式表 = _样式表()
    sst, sst序 = {}, []
    表XML = []
    全部批注 = []              # (sheet序, 批注格, 超链集)
    for 序, tp in enumerate(计划.表):
        w = _表写出(tp, 样式表, sst, sst序)
        表XML.append(w.出())
        if w.批注格 or getattr(w, "超链集", None):
            全部批注.append((序, w.批注格, getattr(w, "超链集", [])))

    # ---- workbook.xml ----
    表片 = []
    for 序, tp in enumerate(计划.表):
        态 = ' state="hidden"' if tp.状态 == "hidden" else ""
        表片.append(f'<sheet name="{_xesc(tp.名称)}" sheetId="{序 + 1}" r:id="rId{序 + 1}"{态}/>')
    定义名 = []
    for 序, tp in enumerate(计划.表):
        p = tp.打印 or {}
        名 = _xesc(tp.名称)
        if p.get("title_rows"):
            定义名.append(f'<definedName name="_xlnm.Print_Titles" localSheetId="{序}">'
                         f"'{名}'!{p['title_rows'].split(':')[0].join(('$', ''))}"
                         f":{p['title_rows'].split(':')[1].join(('$', ''))}</definedName>")
        if p.get("area"):
            定义名.append(f'<definedName name="_xlnm.Print_Area" localSheetId="{序}">'
                          f"'{名}'!{_打印名(p['area'])}</definedName>")
    片 = [_HEADER,
          f"<workbook {_NS}><workbookPr defaultThemeVersion=\"166925\"/>"
          f"<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"20140\" "
          f"windowHeight=\"10960\" activeTab=\"{计划.active}\"/></bookViews><sheets>"]
    片.extend(表片)
    片.append("</sheets>")
    if 定义名:
        片.append(f"<definedNames>{''.join(定义名)}</definedNames>")
    片.append('<calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
    工作簿XML = "".join(片)

    # ---- 打包 ----
    n = len(计划.表)
    with zipfile.ZipFile(输出路径, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as z:
        类型 = [_HEADER,
               '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>'
               '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
               f'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
               f'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
               f'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>']
        for 序 in range(n):
            类型.append(f'<Override PartName="/xl/worksheets/sheet{序 + 1}.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
        for k in range(len(全部批注)):
            类型.append(f'<Override PartName="/xl/comments{k + 1}.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>')
        类型.append("</Types>")
        z.writestr("[Content_Types].xml", "".join(类型).encode("utf-8"))

        z.writestr("_rels/.rels",
                   (_HEADER + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                    "</Relationships>").encode("utf-8"))

        关系 = [_HEADER,
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for 序 in range(n):
            关系.append(f'<Relationship Id="rId{序 + 1}" '
                        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                        f'Target="worksheets/sheet{序 + 1}.xml"/>')
        关系.append(f'<Relationship Id="rId{n + 1}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                    'Target="styles.xml"/>')
        关系.append(f'<Relationship Id="rId{n + 2}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                    'Target="sharedStrings.xml"/>')
        关系.append("</Relationships>")
        z.writestr("xl/_rels/workbook.xml.rels", "".join(关系).encode("utf-8"))

        z.writestr("xl/workbook.xml", 工作簿XML.encode("utf-8"))
        z.writestr("xl/styles.xml", 样式表.xml().encode("utf-8"))
        z.writestr("xl/sharedStrings.xml",
                   (_HEADER + f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                    f'count="{len(sst序)}" uniqueCount="{len(sst序)}">'
                    + "".join(f"<si><t xml:space=\"preserve\">{_xesc(s_)}</t></si>" for s_ in sst序)
                    + "</sst>").encode("utf-8"))

        for 序, x in enumerate(表XML):
            z.writestr(f"xl/worksheets/sheet{序 + 1}.xml", x.encode("utf-8"))
            该表 = next((b for b in 全部批注 if b[0] == 序), None)
            if 该表 is None:
                continue
            _, 批注格, 超链集 = 该表
            条 = []
            if 批注格:
                k = len([1 for b in 全部批注 if b[0] < 序]) + 1
                cx, vml = _批注XML(批注格)
                z.writestr(f"xl/comments{k}.xml", cx.encode("utf-8"))
                z.writestr(f"xl/drawings/vmlDrawing{k}.vml", vml.encode("utf-8"))
                条.append('<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
                          f'Target="../comments{k}.xml"/>')
                条.append('<Relationship Id="rIdLegacy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
                          f'Target="../drawings/vmlDrawing{k}.vml"/>')
            for k, (_, 目标) in enumerate(超链集):
                条.append(f'<Relationship Id="rIdH{k + 1}" '
                          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                          f'Target="{_xesc(目标)}" TargetMode="External"/>')
            z.writestr(
                f"xl/worksheets/_rels/sheet{序 + 1}.xml.rels",
                (_HEADER + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 + "".join(条) + "</Relationships>").encode("utf-8"))
