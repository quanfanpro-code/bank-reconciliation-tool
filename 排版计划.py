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
