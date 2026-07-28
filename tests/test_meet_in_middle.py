"""
折半枚举算法 (Meet-in-the-Middle) 单元测试
"""
import pandas as pd
import pytest

from matcher import _meet_in_middle_solve


# 固定时间戳，供所有测试共用
_NOW = pd.Timestamp("2024-01-15")


def test_三个数组合出目标值():
    """3个数组合出目标值"""
    values = [5000, 4000, 3000]
    dates = [_NOW, _NOW, _NOW]
    indices = [0, 1, 2]
    target = 12000
    result = _meet_in_middle_solve(values, dates, indices, target)
    assert result is not None, "应找到组合"
    result_indices, confidence = result
    assert sorted(result_indices) == [0, 1, 2], f"应返回所有索引，实际: {result_indices}"
    assert confidence in ["高", "中", "低"], f"置信度应为有效值，实际: {confidence}"


def test_五个数组合出目标值():
    """5个数组合出目标值"""
    values = [1000, 2000, 3000, 4000, 5000]
    dates = [_NOW] * 5
    indices = [0, 1, 2, 3, 4]
    target = 10000
    result = _meet_in_middle_solve(values, dates, indices, target)
    assert result is not None, "应找到组合"
    result_indices, confidence = result
    actual_sum = sum(values[i] for i in result_indices if i < len(values))
    assert actual_sum == target, f"组合和应为{target}，实际: {actual_sum}"
    assert len(result_indices) == 3, f"应返回3个元素的组合，实际: {len(result_indices)}个"


def test_无法组合返回None():
    """无法组合出目标值时返回None"""
    values = [1000, 2000, 3000]
    dates = [_NOW] * 3
    indices = [0, 1, 2]
    target = 9999
    result = _meet_in_middle_solve(values, dates, indices, target)
    assert result is None, f"无法组合时应返回None，实际: {result}"


def test_空输入返回None():
    """空输入返回None"""
    result = _meet_in_middle_solve([], [], [], 1000)
    assert result is None, f"空输入应返回None，实际: {result}"


def test_单个数值等于目标值():
    """单个数值等于目标值"""
    values = [5000]
    dates = [_NOW]
    indices = [10]
    target = 5000
    result = _meet_in_middle_solve(values, dates, indices, target)
    assert result is not None, "应找到组合"
    result_indices, confidence = result
    assert result_indices == [10], f"应返回索引[10]，实际: {result_indices}"
    assert confidence == "高", f"单元素应为高置信度，实际: {confidence}"


def test_包含正负数的组合():
    """包含正负数的组合"""
    values = [10000, -2000, 3000, 5000]
    dates = [_NOW] * 4
    indices = [0, 1, 2, 3]
    target = 8000
    result = _meet_in_middle_solve(values, dates, indices, target, allow_mixed_sign=True)
    assert result is not None, "应找到组合"
    result_indices, confidence = result
    assert sorted(result_indices) == [0, 1], (
        f"应返回索引0,1 (10000-2000=8000)，实际: {result_indices}"
    )
