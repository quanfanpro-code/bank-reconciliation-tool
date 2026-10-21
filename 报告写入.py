# -*- coding: utf-8 -*-
"""报告写入：把 排版计划.簿计划 渲染为 xlsx。

写入走 直写XML.py（直接序列化 OpenXML，一次 zip 成型）。
"""
import math

import pandas as pd

from 排版计划 import 簿计划


def _值转换(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if v is pd.NaT:
        return None
    if v == "":
        # 旧路径保存重载后空字符串即 None，样式与取值语义都对齐重载态
        return None
    return v


def 渲染簿计划(计划: 簿计划, 输出路径: str) -> None:
    """渲染簿计划：直写 OpenXML（不经 openpyxl 对象模型），一次 zip 成型。"""
    from 直写XML import 写簿
    写簿(计划, 输出路径)
