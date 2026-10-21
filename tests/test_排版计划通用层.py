# -*- coding: utf-8 -*-
"""通用层回放：底层主题、明细着色、差额着色、期初警告、通用排版。"""
import pandas as pd
import pytest

from 排版计划 import 簿计划, 表计划, 新建簿计划


def _主题表():
    df = pd.DataFrame({
        "名称": ["甲", "乙", "合计"],
        "金额": [1.5, 2.5, 4.0],
        "交易日期": ["2026-01-01", "2026-01-02", "2026-01-03"],
    })
    return 表计划(名称="t", 表=df)


def test_底层主题():
    from 排版计划 import 应用底层主题

    tp = _主题表()
    应用底层主题(tp)
    assert tp.前置空列 == 1
    # 全部行高 18（表头 + 3 数据行）
    assert all(tp.行[r].高 == 18 for r in range(1, 5))
    # 表头样式
    assert tp.表头样式.加粗 and tp.表头样式.填充色 == "1F4E79"
    assert tp.表头样式.水平 == "right" and tp.表头样式.垂直 == "center"
    # 金额列：蓝色右对齐 + 货币格式；名称列黑色左对齐文本
    名称列, 金额列, 日期列 = tp.列[0], tp.列[1], tp.列[2]
    assert 金额列.基础样式.字色 == "0000FF" and 金额列.基础样式.水平 == "right"
    assert 金额列.基础样式.数字格式 == "#,##0.00"
    assert 名称列.基础样式.字色 is None and 名称列.基础样式.水平 == "left"
    assert 日期列.基础样式.数字格式 == "yyyy/mm/dd"
    # 列宽估算（介于 10~50）
    assert all(10 <= c.宽 <= 50 for c in tp.列[:3])
    # 末行"合计"触发合计行样式
    assert tp.行[4].样式.加粗 and tp.行[4].样式.填充色 == "D6E4F0"
    # 边框：首行上粗下虚、末行上虚下粗、中间全虚
    assert tp.表头样式.上框 == ("medium", "000000")
    assert tp.行[2].样式.上框 == ("dashed", "808080")
    assert tp.行[4].样式.下框 == ("medium", "000000")
    # 冻结/筛选/网格线/A 列宽/右侧空列
    assert tp.冻结 == "A2"
    assert tp.筛选 == "B1:D4"
    assert tp.网格线 is False
    assert tp.前部列[0].宽 == 2
    assert tp.列[3].宽 == 3   # 数据列之后的右侧空列


def test_明细着色_删列与风险():
    from 排版计划 import 应用底层主题, 应用明细着色

    df = pd.DataFrame({
        "摘要": ["x", "y"],
        "差额": [0.0, 5.0],
        "风险等级": ["低风险", "高风险"],
        "_低置信度标记": [True, False],
    })
    tp = 表计划(名称="逐笔匹配", 表=df)
    应用底层主题(tp)
    应用明细着色(tp)
    assert "_低置信度标记" not in tp.表.columns
    assert len([c for c in tp.列 if c.宽 is not None]) >= 3  # 列计划同步删
    # 第 2 行（低置信度）整行黄底红粗字
    assert tp.行[2].样式.填充色 == "FFFF00" and tp.行[2].样式.加粗
    assert tp.行[2].样式.字色 == "FF0000"
    # 第 3 行差额 5.0 且高风险 → 差额格 FFC7CE（覆盖优先于行样式）
    差额列号 = list(tp.表.columns).index("差额") + 2
    assert tp.覆盖[(3, 差额列号)].填充色 == "FFC7CE"
    assert (2, 差额列号) not in tp.覆盖   # 差额为 0 不着色


def test_差额着色():
    from 排版计划 import 应用差额着色

    df = pd.DataFrame({"项目": ["a", "b"], "差额": [0.0, -3.5]})
    tp = 表计划(名称="每日统计", 表=df)
    应用差额着色(tp)
    差额列号 = list(tp.表.columns).index("差额") + 2
    assert tp.覆盖[(3, 差额列号)].填充色 == "FFF2CC"
    assert (2, 差额列号) not in tp.覆盖


def test_期初警告():
    from 排版计划 import 应用期初警告

    df = pd.DataFrame({"项目": ["期初余额一致", "其他"], "数值": ["x", "y"]})
    tp = 表计划(名称="核对结论", 表=df)
    应用期初警告(tp, True)
    assert tp.行[2].样式.填充色 == "FFFF00"
    assert tp.行[2].样式.字色 == "FF0000" and tp.行[2].样式.字号 == 14
    assert 3 not in tp.行 or tp.行[3].样式 is None
    # 无警告时不动
    tp2 = 表计划(名称="核对结论", 表=df)
    应用期初警告(tp2, False)
    assert not tp2.行


def test_通用排版():
    from 排版计划 import 应用底层主题, 应用通用排版

    tables = {
        "核对结论": pd.DataFrame({
            "项目": ["待处理事项数", "逐笔精确匹配率"],
            "数值": ["占位", 0.5],
        }),
        "疑点事项": pd.DataFrame({
            "系统结论": ["有待处理"],
            "后续状态": [None],
            "摘要": ["很长的摘要" * 30],
            "匹配ID": ["m1"],
        }),
        "运行参数": pd.DataFrame({"参数": ["x"]}),
    }
    簿 = 新建簿计划(tables)
    for tp in 簿.表:
        应用底层主题(tp)
    应用通用排版(簿)

    结论 = 簿.取表("核对结论")
    疑点 = 簿.取表("疑点事项")
    参数 = 簿.取表("运行参数")

    # 技术表隐藏、A 列隐藏、冻结到首个业务列
    assert 参数.状态 == "hidden"
    assert 疑点.前部列[0].隐藏
    assert 疑点.冻结 == "B2"
    # 摘要列宽 28，其余 14；疑点事项的 匹配ID 隐藏
    摘要下标 = list(疑点.表.columns).index("摘要")
    assert 疑点.列[摘要下标].宽 == 28
    匹配下标 = list(疑点.表.columns).index("匹配ID")
    assert 疑点.列[匹配下标].隐藏 and 疑点.列[匹配下标].宽 == 14
    # 数据格：水平清空 + 顶端 + 换行
    assert 疑点.列[摘要下标].基础样式.水平 == ""
    assert 疑点.列[摘要下标].基础样式.垂直 == "top"
    assert 疑点.列[摘要下标].基础样式.自动换行
    # 长摘要撑高行（按 28 宽估算，大于 18）
    assert 疑点.行[2].高 > 18
    # 疑点事项挂验证
    assert len(疑点.验证) == 1 and 疑点.验证[0]["type"] == "list"
    状态字母 = chr(ord("A") + 1 + list(疑点.表.columns).index("后续状态"))
    assert 疑点.验证[0]["sqref"].startswith(状态字母 + "2:")
    # 核对结论：待处理事项数被改写为 COUNTIFS 公式，百分比行数字格式
    assert str(结论.表.iloc[0, 1]).startswith("=COUNTIFS(")
    数值下标 = list(结论.表.columns).index("数值")
    assert 结论.覆盖[(3, 数值下标 + 2)].数字格式 == "0.00%"
    assert 簿.自动重算
