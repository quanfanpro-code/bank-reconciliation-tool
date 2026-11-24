# -*- coding: utf-8 -*-
"""报告写入：把 排版计划.簿计划 渲染为 xlsx。

写入走 直写XML.py（直接序列化 OpenXML，一次 zip 成型）。
"""
import math

import numpy as np
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
    if v.__class__.__module__ == "numpy":
        # numpy 标量（np.int64 等）不是 int/float 子类，不归一会被直写静默丢弃
        if isinstance(v, np.bool_):
            return bool(v)
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            f = float(v)
            return None if math.isnan(f) or math.isinf(f) else f
        if isinstance(v, np.datetime64):
            return pd.Timestamp(v).to_pydatetime()
        return str(v)
    return v


def 渲染簿计划(计划: 簿计划, 输出路径: str) -> None:
    """渲染簿计划：直写 OpenXML（不经 openpyxl 对象模型），一次 zip 成型。"""
    from 直写XML import 写簿
    写簿(计划, 输出路径)
