# ==========================================
# 核心匹配器模块 (Matcher Module)
# ==========================================
# 包含：
#   - 算法函数（DFS、折半枚举、贪心等）
#   - 并行 DFS 处理函数
#   - 核心匹配引擎（Matcher 类）

import gc
import hashlib
import json
import os
import re
import random
import bisect
import secrets
import time
import unicodedata
from contextvars import ContextVar
from dataclasses import dataclass, replace
from datetime import timedelta
from concurrent.futures import ProcessPoolExecutor, as_completed
from concurrent.futures import TimeoutError as FuturesTimeoutError
from concurrent.futures.process import BrokenProcessPool
from typing import Optional, List, Dict, Any, Callable, Tuple, Set
from decimal import Decimal

import pandas as pd
import numpy as np

# 尝试引入 psutil（可选依赖）
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# 本地模块导入
from precision_engine import PrecisionEngine
from data_structures import (
    BusinessClue, BusinessEvent, DifferencePoolResult, LLMDecisionRecord, MatcherConfig,
    MatchCandidate, OverallControlResult,
    ProcessingStatus,
    WorkerExceptionLogger,
    DFS_CONFIDENCE_HIGH_THRESHOLD, DFS_CONFIDENCE_MEDIUM_THRESHOLD,
    MEET_IN_MIDDLE_THRESHOLD, DATE_OFFSET_PRIORITIES
)
from llm_assistant import (
    ONLINE_ALLOWED_FIELDS,
    CandidateSemanticRequest,
    SemanticCandidate,
)
from matching_policy import (
    _critical_field_category,
    apply_monthly_difference_pools,
    build_sensitive_field_signals,
    build_group_metrics,
    bucket_distribution,
    merge_labeled_fields,
    route_candidate,
    score_candidate,
    score_text_fields,
)
from utils import normalize_summary
from 业务分组 import (
    row_business, business_evidence, complete_groups, has_business_conflict,
    relationship_priority, candidate_sort_key,
)
from 业务事件 import business_identity, detect_same_side_events, has_fee_evidence, rows_share_business

# ==========================================
# 辅助函数
# ==========================================

def _gc_cleanup(stage_name: str = "", logger: Optional[Callable[[str], None]] = None) -> None:
    """执行垃圾回收并记录内存使用

    :param stage_name: 当前阶段名称
    :param logger: 日志函数
    """
    collected = gc.collect()
    if logger and PSUTIL_AVAILABLE:
        try:
            mem_gb = psutil.Process().memory_info().rss / (1024 ** 3)
            logger(f"🧹 [{stage_name}] GC回收 {collected} 个对象，当前内存: {mem_gb:.2f}GB")
        except (OSError, RuntimeError):
            pass


# ==========================================
# 算法函数
# ==========================================


@dataclass
class _CombinationSearchBudget:
    node_limit: int
    deadline: float
    nodes_visited: int = 0
    exhaustion_reason: str = ""

    def allows_work(self, consume_node: bool = True) -> bool:
        if self.exhaustion_reason:
            return False
        if time.monotonic() >= self.deadline:
            self.exhaustion_reason = "task_timeout"
            return False
        if consume_node:
            if self.nodes_visited >= max(0, int(self.node_limit)):
                self.exhaustion_reason = "node_limit"
                return False
            self.nodes_visited += 1
        return True


@dataclass
class _CombinationTaskResult:
    source_idx: int
    legacy_result: Optional[Tuple[int, List[List[int]], str]]
    candidate_count: int
    retained_count: int
    depth_limited: bool
    nodes_visited: int
    exhaustion_reason: str = ""

    @property
    def candidate_truncated(self) -> bool:
        return self.candidate_count > self.retained_count

    @property
    def fully_searched(self) -> bool:
        return not (
            self.candidate_truncated
            or self.depth_limited
            or self.exhaustion_reason
        )


_ACTIVE_COMBINATION_BUDGET: ContextVar[Optional[_CombinationSearchBudget]] = (
    ContextVar("active_combination_budget", default=None)
)


def _combination_budget_allows_work(consume_node: bool = True) -> bool:
    budget = _ACTIVE_COMBINATION_BUDGET.get()
    return budget is None or budget.allows_work(consume_node)

def _dfs_solve_static(values: List[int], dates: List[pd.Timestamp], indices: List[int],
                    target: int, max_depth: int, allow_mixed_sign: bool = False,
                    date_window: int = 3) -> Optional[Tuple[List[int], str]]:
    """
    静态版本的 DFS 求解子集和问题（用于多进程）
    """
    n = len(values)
    if n == 0:
        return None

    suffix_sum = [0] * (n + 1)
    for i in range(n - 1, -1, -1):
        suffix_sum[i] = suffix_sum[i + 1] + values[i]

    solution_info = None

    def backtrack(start_idx, current_sum, path_idxs, path_dates):
        nonlocal solution_info
        if not _combination_budget_allows_work():
            return
        if solution_info is not None:
            return

        if current_sum == target:
            if len(path_idxs) <= DFS_CONFIDENCE_HIGH_THRESHOLD:
                confidence = '高'
            elif len(path_idxs) <= DFS_CONFIDENCE_MEDIUM_THRESHOLD:
                confidence = '中'
            else:
                confidence = '低'

            if len(path_dates) > 1:
                ts_dates = [pd.Timestamp(d) for d in path_dates]
                date_span = max(ts_dates) - min(ts_dates)
                if date_span > timedelta(days=date_window):
                    if confidence == '高':
                        confidence = '中'
                    elif confidence == '中':
                        confidence = '低'

            solution_info = (list(path_idxs), confidence)
            return

        if len(path_idxs) >= max_depth:
            return

        if not allow_mixed_sign:
            if target > 0:
                if current_sum > target: return
            else:
                if current_sum < target: return

        if not allow_mixed_sign:
            if target > 0:
                if current_sum + suffix_sum[start_idx] < target: return
            else:
                if current_sum + suffix_sum[start_idx] > target: return

        for i in range(start_idx, n):
            val = values[i]
            if i > start_idx and values[i] == values[i-1]:
                continue

            path_idxs.append(indices[i])
            path_dates.append(dates[i])
            backtrack(i + 1, current_sum + val, path_idxs, path_dates)
            path_dates.pop()
            path_idxs.pop()

    backtrack(0, 0, [], [])
    return solution_info


def _meet_in_middle_solve(values: List[int], dates: List[pd.Timestamp], indices: List[int],
                          target: int, max_depth: int = 30, allow_mixed_sign: bool = False,
                          date_window: int = 31) -> Optional[Tuple[List[int], str]]:
    """
    折半枚举算法 (Meet-in-the-Middle) 求解子集和问题。
    
    时间复杂度: O(2^(N/2)) 相比 DFS 的 O(2^N)
    
    算法原理:
    1. 将数组从中间劈成两半 (left, right)
    2. 对左半边用 DFS 计算所有子集和，存入字典 {sum: [(indices, dates), ...]}
    3. 对右半边也计算所有子集和，每算一个 sum_right，在左字典中查找 target - sum_right
    4. 找到则合并结果，计算置信度
    5. 优先返回元素数最少的组合
    """
    n = len(values)
    if n == 0:
        return None
    
    if n == 1:
        if values[0] == target:
            confidence = '高'
            return ([indices[0]], confidence)
        return None
    
    mid = n // 2
    left_values = values[:mid]
    left_dates = dates[:mid]
    left_indices = indices[:mid]
    
    right_values = values[mid:]
    right_dates = dates[mid:]
    right_indices = indices[mid:]
    
    left_subsets = {}
    
    def dfs_left(start_idx, current_sum, path_idxs, path_dates):
        if not _combination_budget_allows_work():
            return
        if len(path_idxs) > max_depth:
            return
        
        key = current_sum
        # 同一合计只保留笔数最少的一解，避免同额记录产生笛卡尔展开。
        existing = left_subsets.get(key)
        if existing is None or (len(path_idxs), tuple(path_idxs)) < (len(existing[0]), tuple(existing[0])):
            left_subsets[key] = (list(path_idxs), list(path_dates))
        for i in range(start_idx, len(left_values)):
            val = left_values[i]
            if not allow_mixed_sign:
                if target > 0 and current_sum + val > target:
                    continue
            path_idxs.append(left_indices[i])
            path_dates.append(left_dates[i])
            dfs_left(i + 1, current_sum + val, path_idxs, path_dates)
            path_dates.pop()
            path_idxs.pop()
    
    dfs_left(0, 0, [], [])
    
    best_solution = None
    
    def dfs_right(start_idx, current_sum, path_idxs, path_dates):
        nonlocal best_solution
        if not _combination_budget_allows_work():
            return
        if len(path_idxs) > max_depth:
            return
        
        needed = target - current_sum
        if needed in left_subsets:
            left_idxs, left_dts = left_subsets[needed]
            combined_idxs = left_idxs + path_idxs
            combined_dates = left_dts + path_dates
            if len(combined_idxs) <= max_depth and (
                best_solution is None or
                (len(combined_idxs), tuple(combined_idxs)) <
                (len(best_solution[0]), tuple(best_solution[0]))
            ):
                best_solution = (combined_idxs, combined_dates)
        
        for i in range(start_idx, len(right_values)):
            val = right_values[i]
            if not allow_mixed_sign:
                if target > 0 and current_sum + val > target:
                    continue
            path_idxs.append(right_indices[i])
            path_dates.append(right_dates[i])
            dfs_right(i + 1, current_sum + val, path_idxs, path_dates)
            path_dates.pop()
            path_idxs.pop()
    
    dfs_right(0, 0, [], [])
    
    if best_solution is None:
        return None

    best_idxs, best_dates = best_solution
    
    if len(best_idxs) <= DFS_CONFIDENCE_HIGH_THRESHOLD:
        confidence = '高'
    elif len(best_idxs) <= DFS_CONFIDENCE_MEDIUM_THRESHOLD:
        confidence = '中'
    else:
        confidence = '低'
    
    if len(best_dates) > 1:
        ts_dates = [pd.Timestamp(d) for d in best_dates]
        date_span = max(ts_dates) - min(ts_dates)
        if date_span > timedelta(days=date_window):
            if confidence == '高':
                confidence = '中'
            elif confidence == '中':
                confidence = '低'
    
    return (best_idxs, confidence)


def _solve_combination(values: List[int], dates: List[pd.Timestamp], indices: List[int],
                       target: int, max_depth: int = 30, allow_mixed_sign: bool = False,
                       date_window: int = 31) -> Optional[Tuple[List[int], str]]:
    """
    智能选择组合匹配算法。
    
    根据候选数量自动选择最优算法：
    - 候选数 <= MEET_IN_MIDDLE_THRESHOLD: 使用DFS（对小规模数据更高效）
    - 候选数 > MEET_IN_MIDDLE_THRESHOLD: 使用折半枚举（对大规模数据更高效）
    """
    n = len(values)
    if n <= MEET_IN_MIDDLE_THRESHOLD:
        return _dfs_solve_static(values, dates, indices, target, max_depth, allow_mixed_sign, date_window)
    else:
        return _meet_in_middle_solve(values, dates, indices, target, max_depth, allow_mixed_sign, date_window)


def _near_combination_solve(
    values: List[int],
    dates: List[pd.Timestamp],
    indices: List[int],
    target: int,
    tolerance: int,
    max_depth: int = 30,
    allow_mixed_sign: bool = False,
    date_window: int = 31,
) -> Optional[Tuple[List[int], str]]:
    """在容差范围内寻找差额最小、笔数最少且结果稳定的组合。"""
    if tolerance <= 0 or len(values) < 2 or max_depth < 2:
        return None

    normalized_values = [int(value) for value in values]
    normalized_dates = [pd.Timestamp(date) for date in dates]
    normalized_indices = [int(index) for index in indices]
    target = int(target)
    tolerance = int(tolerance)
    midpoint = len(normalized_values) // 2
    def enumerate_half(
        half_values: List[int],
        half_dates: List[pd.Timestamp],
        half_indices: List[int],
    ) -> Optional[List[Dict[int, Tuple[Tuple[int, ...], Tuple[pd.Timestamp, ...]]]]]:
        by_count: List[
            Dict[int, Tuple[Tuple[int, ...], Tuple[pd.Timestamp, ...]]]
        ] = [dict() for _ in range(min(max_depth, len(half_values)) + 1)]
        def walk(
            position: int,
            current_sum: int,
            path_indices: List[int],
            path_dates: List[pd.Timestamp],
        ) -> None:
            if not _combination_budget_allows_work():
                return

            count = len(path_indices)
            if count <= max_depth:
                key = tuple(sorted(path_indices))
                existing = by_count[count].get(current_sum)
                if existing is None or key < tuple(sorted(existing[0])):
                    by_count[count][current_sum] = (
                        tuple(path_indices),
                        tuple(path_dates),
                    )
            if position >= len(half_values) or count >= max_depth:
                return

            for item_position in range(position, len(half_values)):
                next_sum = current_sum + half_values[item_position]
                if not allow_mixed_sign:
                    if target > 0 and next_sum > target + tolerance:
                        continue
                    if target < 0 and next_sum < target - tolerance:
                        continue
                path_indices.append(half_indices[item_position])
                path_dates.append(half_dates[item_position])
                walk(
                    item_position + 1,
                    next_sum,
                    path_indices,
                    path_dates,
                )
                path_dates.pop()
                path_indices.pop()

        walk(0, 0, [], [])
        return by_count

    left_by_count = enumerate_half(
        normalized_values[:midpoint],
        normalized_dates[:midpoint],
        normalized_indices[:midpoint],
    )
    right_by_count = enumerate_half(
        normalized_values[midpoint:],
        normalized_dates[midpoint:],
        normalized_indices[midpoint:],
    )
    right_sums_by_count = [
        sorted(sum_map)
        for sum_map in right_by_count
    ]
    best: Optional[
        Tuple[
            Tuple[int, int, Tuple[int, ...]],
            Tuple[int, ...],
            Tuple[pd.Timestamp, ...],
        ]
    ] = None

    for left_count, left_sum_map in enumerate(left_by_count):
        if not _combination_budget_allows_work(False):
            break
        for left_sum, (left_indices, left_dates) in left_sum_map.items():
            if not _combination_budget_allows_work():
                break
            minimum_right_count = max(0, 2 - left_count)
            maximum_right_count = min(
                len(right_by_count) - 1,
                max_depth - left_count,
            )
            for right_count in range(
                minimum_right_count,
                maximum_right_count + 1,
            ):
                if not _combination_budget_allows_work():
                    break
                sorted_right_sums = right_sums_by_count[right_count]
                if not sorted_right_sums:
                    continue
                needed = target - left_sum
                insertion = bisect.bisect_left(sorted_right_sums, needed)
                for right_position in (insertion - 1, insertion):
                    if not 0 <= right_position < len(sorted_right_sums):
                        continue
                    right_sum = sorted_right_sums[right_position]
                    difference = abs(target - left_sum - right_sum)
                    if difference > tolerance:
                        continue
                    right_indices, right_dates = right_by_count[
                        right_count
                    ][right_sum]
                    combined_indices = tuple(
                        sorted(left_indices + right_indices)
                    )
                    ranking = (
                        difference,
                        len(combined_indices),
                        combined_indices,
                    )
                    combined_dates = left_dates + right_dates
                    if best is None or ranking < best[0]:
                        best = (
                            ranking,
                            combined_indices,
                            combined_dates,
                        )

    if best is None:
        return None

    best_indices = list(best[1])
    best_dates = best[2]
    if len(best_indices) <= DFS_CONFIDENCE_HIGH_THRESHOLD:
        confidence = "高"
    elif len(best_indices) <= DFS_CONFIDENCE_MEDIUM_THRESHOLD:
        confidence = "中"
    else:
        confidence = "低"

    if len(best_dates) > 1:
        date_span = max(best_dates) - min(best_dates)
        if date_span > timedelta(days=date_window):
            if confidence == "高":
                confidence = "中"
            elif confidence == "中":
                confidence = "低"

    return best_indices, confidence


def _randomized_greedy(window_amounts: List[int], window_dates: List[pd.Timestamp],
                       window_indices: List[int], target: int,
                       num_attempts: int = 3, random_seed: int = 0) -> Optional[Tuple[List[int], str]]:
    """
    随机化贪心策略（可控随机，使用固定种子确保同数据同结果）
    """
    rng = random.Random(random_seed)
    
    if not window_amounts:
        return None

    # DataFrame 为节省内存可能把正数金额压缩成 uint32。
    # 先恢复为 Python int，避免目标额减去较大候选时发生无符号回绕。
    target = int(target)
    candidates = [
        (int(amount), date, int(index))
        for amount, date, index in zip(
            window_amounts,
            window_dates,
            window_indices,
        )
    ]
    for attempt in range(num_attempts):
        if not _combination_budget_allows_work():
            return None
        if attempt == 0:
            # 第 0 次：确定性排序基线（大金额优先）
            trial = sorted(candidates, key=lambda x: abs(x[0]), reverse=True)
        else:
            # 后续尝试：真正洗牌（修复前先 shuffle 后 sort，随机性被中和）
            trial = candidates.copy()
            rng.shuffle(trial)

        result_indices = []
        current_sum = 0
        current_diff = abs(target)

        for amount, date, idx in trial:
            if not _combination_budget_allows_work():
                return None
            new_sum = current_sum + amount
            new_diff = abs(target - new_sum)

            if new_diff < current_diff:
                result_indices.append(idx)
                current_sum = new_sum
                current_diff = new_diff

            if current_sum == target:
                break

        if current_sum == target:
            confidence = '高' if len(result_indices) <= 2 else ('中' if len(result_indices) <= 5 else '低')
            return (result_indices, confidence)

    return None


def _process_single_source_with_diagnostics(args: Tuple) -> _CombinationTaskResult:
    """处理一个组合来源，并返回匹配结果及搜索预算使用情况。"""
    source_idx, source_date, target_val, targets_data, config = args[:5]
    timeout_seconds = (
        float(args[5])
        if len(args) > 5
        else float(config.combination_task_timeout_seconds)
    )
    target_val = int(target_val)
    max_depth = config.max_dfs_depth
    allow_mixed_sign = config.allow_mixed_sign
    max_candidates = min(config.max_candidates, 30)
    date_window = config.dfs_date_window
    budget = _CombinationSearchBudget(
        node_limit=config.combination_node_limit_per_source,
        deadline=time.monotonic() + max(0.0, timeout_seconds),
    )
    budget_token = _ACTIVE_COMBINATION_BUDGET.set(budget)
    candidate_count = 0
    retained_count = 0
    depth_limited = False

    def finish(
        result: Optional[Tuple[int, List[List[int]], str]] = None,
    ) -> _CombinationTaskResult:
        return _CombinationTaskResult(
            source_idx=int(source_idx),
            legacy_result=result,
            candidate_count=candidate_count,
            retained_count=retained_count,
            depth_limited=depth_limited,
            nodes_visited=budget.nodes_visited,
            exhaustion_reason=budget.exhaustion_reason,
        )

    try:
        window_indices: List[int] = []
        window_amounts: List[int] = []
        window_dates: List[Any] = []
        window_stable_keys: List[str] = []
        for item in targets_data['view_dict']:
            if item['matched']:
                continue
            amount_decimal = int(item['amount_decimal'])
            if not allow_mixed_sign and target_val != 0:
                if target_val > 0 and amount_decimal <= 0:
                    continue
                if target_val < 0 and amount_decimal >= 0:
                    continue
            window_indices.append(int(item['index']))
            window_amounts.append(amount_decimal)
            window_dates.append(item['date'])
            window_stable_keys.append(str(item.get("stable_key", "")))

        if target_val == 0:
            candidate_count = len(window_indices)
            retained_count = min(candidate_count, max_candidates)
            if candidate_count < 2:
                return finish()
            if candidate_count > max_candidates:
                indexed = sorted(
                    zip(window_amounts, window_dates, window_indices, window_stable_keys),
                    key=lambda item: (-abs(item[0]), item[3], item[2]),
                )[:max_candidates]
                window_amounts = [item[0] for item in indexed]
                window_dates = [item[1] for item in indexed]
                window_indices = [item[2] for item in indexed]
                window_stable_keys = [item[3] for item in indexed]
            if not config.allow_zero_match:
                return finish()
            window_amounts_int = [
                PrecisionEngine.to_integer_li(amount)
                for amount in window_amounts
            ]
            if 0 in window_amounts_int:
                zero_pos = window_amounts_int.index(0)
                return finish(
                    (int(source_idx), [[window_indices[zero_pos]]], '中')
                )
            return finish()

        indexed = sorted(
            zip(window_amounts, window_dates, window_indices, window_stable_keys),
            key=lambda item: (-abs(item[0]), item[3], item[2]),
        )
        window_amounts = [item[0] for item in indexed]
        window_dates = [item[1] for item in indexed]
        window_indices = [item[2] for item in indexed]
        window_stable_keys = [item[3] for item in indexed]

        trivial_tolerance_li = PrecisionEngine.to_integer_li(
            config.clearly_trivial_threshold
        )
        exact_single_exists = any(
            int(amount) == target_val for amount in window_amounts
        )
        if exact_single_exists:
            multi_items = [
                (amount, date, index)
                for amount, date, index in zip(
                    window_amounts,
                    window_dates,
                    window_indices,
                )
                if int(amount) != target_val
            ]
            candidate_count = len(multi_items)
            retained_count = min(candidate_count, max_candidates)
            if candidate_count > max_candidates:
                multi_items = multi_items[:max_candidates]
            result_info = None
            if len(multi_items) >= 2:
                remaining_max = max(abs(int(item[0])) for item in multi_items)
                remaining_depth = min(max_depth, 8)
                if remaining_max > 0:
                    remaining_depth = min(
                        remaining_depth,
                        2 + int(abs(target_val) / remaining_max),
                    )
                depth_limited = remaining_depth < len(multi_items)
                result_info = _solve_combination(
                    [item[0] for item in multi_items],
                    [item[1] for item in multi_items],
                    [item[2] for item in multi_items],
                    target_val,
                    remaining_depth,
                    allow_mixed_sign,
                    date_window,
                )
            if result_info and len(result_info[0]) >= 2:
                return finish(
                    (int(source_idx), [result_info[0]], result_info[1])
                )
            return finish()

        candidate_count = len(window_indices)
        retained_count = min(candidate_count, max_candidates)
        if candidate_count < 2:
            return finish()
        if candidate_count > max_candidates:
            window_amounts = window_amounts[:max_candidates]
            window_dates = window_dates[:max_candidates]
            window_indices = window_indices[:max_candidates]

        max_abs_cand = max(abs(amount) for amount in window_amounts)
        curr_max_depth = max_depth
        if max_abs_cand > 0:
            curr_max_depth = min(
                max_depth,
                2 + int(abs(target_val) / max_abs_cand),
            )
        depth_limited = curr_max_depth < len(window_indices)

        result_info = _solve_combination(
            list(window_amounts),
            list(window_dates),
            list(window_indices),
            target_val,
            curr_max_depth,
            allow_mixed_sign,
            date_window,
        )
        if (
            result_info is None
            and trivial_tolerance_li > 0
            and not budget.exhaustion_reason
        ):
            result_info = _near_combination_solve(
                list(window_amounts),
                list(window_dates),
                list(window_indices),
                target_val,
                trivial_tolerance_li,
                curr_max_depth,
                allow_mixed_sign,
                date_window,
            )

        if result_info:
            result_idxs, confidence = result_info
            solutions = [result_idxs]
            for excluded_index in result_idxs:
                if budget.exhaustion_reason:
                    break
                remaining = [
                    (amount, date, index)
                    for amount, date, index in zip(
                        window_amounts,
                        window_dates,
                        window_indices,
                    )
                    if index != excluded_index
                ]
                if len(remaining) < 2:
                    continue
                alternative = _solve_combination(
                    [item[0] for item in remaining],
                    [item[1] for item in remaining],
                    [item[2] for item in remaining],
                    target_val,
                    curr_max_depth,
                    allow_mixed_sign,
                    date_window,
                )
                if alternative and len(alternative[0]) >= 2:
                    solutions.append(alternative[0])
                    break
            return finish((int(source_idx), solutions, confidence))

        if config.allow_greedy_fallback and not budget.exhaustion_reason:
            greedy_result = _randomized_greedy(
                list(window_amounts),
                list(window_dates),
                list(window_indices),
                target_val,
                config.greedy_attempts,
                config.random_seed,
            )
            if greedy_result and len(greedy_result[0]) >= 2:
                result_idxs, confidence = greedy_result
                return finish(
                    (int(source_idx), [result_idxs], confidence)
                )
        return finish()
    finally:
        _ACTIVE_COMBINATION_BUDGET.reset(budget_token)


def _process_single_source(args: Tuple) -> Optional[Tuple[int, List[List[int]], str]]:
    """兼容旧调用：仅返回组合匹配结果。"""
    return _process_single_source_with_diagnostics(args).legacy_result


# ==========================================
# 核心匹配引擎
# ==========================================


def select_non_conflicting_candidates(
    candidates: List[MatchCandidate],
    *,
    diagnostics: Optional[Dict[str, Any]] = None,
    should_stop: Optional[Callable[[], bool]] = None,
    exact_component_limit: int = 18,
) -> List[MatchCandidate]:
    """按共享源行分组，在小组内整体寻找最优的不重叠关系组合。"""
    stats: Dict[str, Any] = {
        "component_count": 0,
        "exact_components": 0,
        "fallback_components": 0,
        "largest_component_candidates": 0,
        "stopped": False,
        "search_fully_exhausted": True,
        "exact_component_limit": int(exact_component_limit),
    }
    eligible = [
        candidate for candidate in candidates
        if not candidate.evidence.get("selection_ineligible")
        and not (
            has_business_conflict(candidate)
            and candidate.metrics.total_diff_li
            and not candidate.evidence.get("atomic_voucher_group")
        )
    ]

    def components(items: List[MatchCandidate]) -> List[List[MatchCandidate]]:
        if not items:
            return []
        parent = list(range(len(items)))

        def find(position: int) -> int:
            while parent[position] != position:
                parent[position] = parent[parent[position]]
                position = parent[position]
            return position

        def union(left: int, right: int) -> None:
            left_root, right_root = find(left), find(right)
            if left_root != right_root:
                parent[right_root] = left_root

        bank_owner: Dict[int, int] = {}
        journal_owner: Dict[int, int] = {}
        for position, candidate in enumerate(items):
            for index in candidate.bank_idxs:
                union(position, bank_owner.setdefault(int(index), position))
            for index in candidate.journal_idxs:
                union(position, journal_owner.setdefault(int(index), position))
        grouped: Dict[int, List[MatchCandidate]] = {}
        for position, candidate in enumerate(items):
            grouped.setdefault(find(position), []).append(candidate)
        return sorted(
            grouped.values(),
            key=lambda group: min(item.candidate_id for item in group),
        )

    def objective(items: List[MatchCandidate]) -> tuple[Any, ...]:
        full_strong_rows = sum(
            len(item.bank_idxs) + len(item.journal_idxs)
            for item in items
            if item.evidence.get("resolves_full_group")
            and (
                item.evidence.get("complete_business_id")
                or item.evidence.get("complete_business_group")
                or item.evidence.get("shared_transaction_id")
                or item.evidence.get("salary_group")
                or item.match_type == "fee_net"
            )
        )
        strong_rows = sum(
            (len(item.bank_idxs) + len(item.journal_idxs))
            * min(4, int(item.evidence.get("business_strength", 0)))
            for item in items
        )
        covered_rows = sum(len(item.bank_idxs) + len(item.journal_idxs) for item in items)
        zero_diff_rows = sum(
            len(item.bank_idxs) + len(item.journal_idxs)
            for item in items
            if item.metrics.total_diff_li == 0
        )
        unambiguous_rows = sum(
            len(item.bank_idxs) + len(item.journal_idxs)
            for item in items
            if not item.is_ambiguous
        )
        return (
            full_strong_rows,
            int(bool(items)),
            unambiguous_rows,
            zero_diff_rows,
            covered_rows,
            strong_rows,
            sum(item.scores.total for item in items),
            -sum(item.metrics.total_diff_li for item in items),
            -len(items),
        )

    def greedy(group: List[MatchCandidate]) -> List[MatchCandidate]:
        chosen: List[MatchCandidate] = []
        used_bank: Set[int] = set()
        used_journal: Set[int] = set()
        for candidate in sorted(group, key=candidate_sort_key):
            if used_bank.intersection(candidate.bank_idxs) or used_journal.intersection(candidate.journal_idxs):
                continue
            chosen.append(candidate)
            used_bank.update(candidate.bank_idxs)
            used_journal.update(candidate.journal_idxs)
        return chosen

    selected: List[MatchCandidate] = []
    groups = components(eligible)
    stats["component_count"] = len(groups)
    for group in groups:
        stats["largest_component_candidates"] = max(stats["largest_component_candidates"], len(group))
        if should_stop and should_stop():
            stats["stopped"] = True
            stats["search_fully_exhausted"] = False
            selected.extend(greedy(group))
            stats["fallback_components"] += 1
            continue
        if len(group) > max(1, int(exact_component_limit)):
            selected.extend(greedy(group))
            stats["fallback_components"] += 1
            stats["search_fully_exhausted"] = False
            continue
        ordered = sorted(group, key=candidate_sort_key)
        best: List[MatchCandidate] = []
        best_objective = objective(best)

        def stable_choice_key(items: List[MatchCandidate]) -> tuple[Any, ...]:
            return tuple(candidate_sort_key(item) for item in sorted(items, key=candidate_sort_key))

        def search(position: int, chosen: List[MatchCandidate], used_bank: Set[int], used_journal: Set[int]) -> None:
            nonlocal best, best_objective
            if should_stop and should_stop():
                stats["stopped"] = True
                return
            if position >= len(ordered):
                score = objective(chosen)
                if score > best_objective or (
                    score == best_objective
                    and (not best or stable_choice_key(chosen) < stable_choice_key(best))
                ):
                    best, best_objective = list(chosen), score
                return
            candidate = ordered[position]
            if not used_bank.intersection(candidate.bank_idxs) and not used_journal.intersection(candidate.journal_idxs):
                search(
                    position + 1,
                    [*chosen, candidate],
                    used_bank.union(candidate.bank_idxs),
                    used_journal.union(candidate.journal_idxs),
                )
            search(position + 1, chosen, used_bank, used_journal)

        search(0, [], set(), set())
        if stats["stopped"]:
            stats["search_fully_exhausted"] = False
            selected.extend(greedy(group))
            stats["fallback_components"] += 1
        else:
            selected.extend(best)
            stats["exact_components"] += 1
    if diagnostics is not None:
        diagnostics.clear()
        diagnostics.update(stats)
    return sorted(selected, key=candidate_sort_key)


class Matcher:
    MAX_ALTERNATIVE_REFERENCES = 10

    MATCHING_RULES = {
        'strong_rules': {
            '手续费': ['手续费', '账户管理费', '年费', '网银费', '服务费', '费率', '扣费', '汇费', '转账费'],
            '利息': ['利息', '活期利息', '存款利息', '结息', '利息收入', '利息支出', '理财收益'],
        },
        'positive_rules': {
            '工资': ['工资', '代发工资', '薪资', '薪酬', '薪金', '代发', '发放', '津贴', '奖金'],
            '税费': ['税费', '印花税', '增值税', '附加税', '所得税', '个税', '税款', '税务', '代扣代缴'],
        },
        'negative_rules': {
            '泛词': ['往来款', '货款', '转账', '汇款', '备用金', '借款', '还款'],
        },
        'weak_rules': {
            '收款': ['收款', '回款', '入账'],
        }
    }
    
    WHITELIST_RULES = {
        '手续费': ['手续费', '服务费', '费率', '扣费', '汇费', '转账费'],
        '利息': ['利息', '结息', '利息收入', '利息支出', '理财收益', '存款利息'],
        '工资': ['工资', '薪酬', '薪金', '代发', '发放', '津贴', '奖金'],
        '税费': ['增值税', '附加税', '所得税', '个税', '税款', '税务', '代扣代缴', '印花税']
    }
    
    DISTRIBUTION_DIFF_RATIO = 0.3
    DAILY_MAX_COUNT_DIFF_BASE = 3
    DAILY_MAX_COUNT_DIFF_RATIO = 0.05
    MONTHLY_MAX_COUNT_DIFF_BASE = 5
    MONTHLY_MAX_COUNT_DIFF_RATIO = 0.1
    PARALLEL_MIN_TASKS = 64
    PARALLEL_TASKS_PER_WORKER = 4

    def __init__(self, bank_df: pd.DataFrame, journal_df: pd.DataFrame,
                 config: MatcherConfig,
                 logger: Optional[Callable[[str], None]] = None,
                 progress_callback: Optional[Callable[[float], None]] = None,
                 whitelist_rules: Optional[Dict[str, List[str]]] = None,
                 llm_assistant: Any = None,
                 overall_control: Optional[OverallControlResult] = None):
        self.bank = bank_df.copy()
        self.journal = journal_df.copy()
        self.bank['date'] = pd.to_datetime(self.bank['date'])
        self.journal['date'] = pd.to_datetime(self.journal['date'])
        
        self.config = config
        self.logger = logger
        self.progress_callback = progress_callback
        self.exception_logger = WorkerExceptionLogger()
        
        self.bank['matched'] = False
        self.bank['match_id'] = None
        self.bank['match_type'] = None
        self.bank['confidence'] = None
        self.bank['confidence_score'] = 0
        self.bank['processing_status'] = ProcessingStatus.NO_CANDIDATE.value
        
        self.journal['matched'] = False
        self.journal['match_id'] = None
        self.journal['match_type'] = None
        self.journal['confidence'] = None
        self.journal['confidence_score'] = 0
        self.journal['processing_status'] = ProcessingStatus.NO_CANDIDATE.value
        
        self.matches: List[Dict[str, Any]] = []
        self.candidates: List[MatchCandidate] = []
        self.selected_candidates: List[MatchCandidate] = []
        self.llm_records: List[LLMDecisionRecord] = []
        self.difference_pools: List[DifferencePoolResult] = []
        self.business_events: List[BusinessEvent] = []
        self.business_clues: List[BusinessClue] = []
        self._candidate_ids: Set[str] = set()
        self._candidate_index_keys: Set[str] = set()
        self._stable_id_counts: Dict[str, int] = {}
        self.whitelist_rules = whitelist_rules if whitelist_rules else self.WHITELIST_RULES
        self.llm_assistant = llm_assistant
        self.overall_control = overall_control
        self.overall_scope_limited = bool(
            overall_control is not None and overall_control.scope_limited
        )
        self.balance_integrity_limited = bool(
            overall_control is not None
            and (
                overall_control.bank_balance_status == "疑点"
                or overall_control.journal_balance_status == "疑点"
                or bool(overall_control.continuity_anomalies)
                or any(
                    marker in str(reason)
                    for reason in overall_control.reasons
                    for marker in ("余额", "期初", "期末", "连续性")
                )
            )
        )
        requested_seed = self.config.random_seed
        actual_seed = (
            secrets.randbits(64) if requested_seed == -1 else requested_seed
        )
        self.run_parameters = {
            "requested_random_seed": requested_seed,
            "actual_random_seed": actual_seed,
            "candidate_search": {},
        }
        self._candidate_search_truncated_sources: Dict[str, Set[Any]] = {}
        self._row_content_hash_cache: Dict[str, Dict[int, str]] = {
            "bank": {},
            "journal": {},
        }
        self._execution_config = replace(self.config, random_seed=actual_seed)
        self._collecting_candidates = False
        self.stopping = False
        self._business_rows = {
            "bank": {int(index): row_business(row) for index, row in self.bank.iterrows()},
            "journal": {int(index): row_business(row) for index, row in self.journal.iterrows()},
        }
        self._global_profile_indexes: Dict[
            str, Dict[Tuple[Any, ...], Set[int]]
        ] = {"bank": {}, "journal": {}}
        for side, profiles in self._business_rows.items():
            side_index = self._global_profile_indexes[side]
            for index, profile in profiles.items():
                for profile_key in self._profile_pool_keys(profile):
                    side_index.setdefault(profile_key, set()).add(int(index))
        self._atomic_groups: Dict[str, List[frozenset[int]]] = {
            "bank": [],
            "journal": [],
        }

    def match_special_business_events(self) -> None:
        """识别同侧业务链、重复线索和有明确费用证据的净额关系。"""
        self.business_events = []
        self.business_clues = []
        for source, frame in (("bank", self.bank), ("journal", self.journal)):
            events, clues = detect_same_side_events(
                frame,
                source,
                self.config.dfs_date_window,
            )
            self.business_events.extend(events)
            self.business_clues.extend(clues)

        self._reset_candidate_search_stage("special_business")
        for left_name, left, right_name, right in (
            ("bank", self.bank, "journal", self.journal),
            ("journal", self.journal, "bank", self.bank),
        ):
            fee_indexes = [int(index) for index, row in left.iterrows() if has_fee_evidence(row)]
            fee_set = set(fee_indexes)
            main_indexes = [int(index) for index in left.index if int(index) not in fee_set]
            main_by_id: Dict[str, List[int]] = {}
            main_by_party: Dict[str, List[int]] = {}
            for index in main_indexes:
                business_id, party = business_identity(left.loc[index])
                if business_id:
                    main_by_id.setdefault(business_id, []).append(index)
                if party:
                    main_by_party.setdefault(party, []).append(index)
            right_by_amount: Dict[int, List[int]] = {}
            for index, row in right.iterrows():
                right_by_amount.setdefault(int(row["amount_decimal"]), []).append(int(index))
            for fee_idx in fee_indexes:
                fee_row = left.loc[fee_idx]
                fee_amount = int(fee_row["amount_decimal"])
                fee_business_id, fee_party = business_identity(fee_row)
                scoped_main_indexes = set(main_by_id.get(fee_business_id, ())) if fee_business_id else set()
                if fee_party:
                    scoped_main_indexes.update(main_by_party.get(fee_party, ()))
                for main_idx in sorted(scoped_main_indexes):
                    main_row = left.loc[main_idx]
                    if fee_amount * int(main_row["amount_decimal"]) >= 0:
                        continue
                    if abs((pd.Timestamp(main_row["date"]) - pd.Timestamp(fee_row["date"])).days) > self.config.tolerance_days:
                        continue
                    shared_fee, fee_basis = rows_share_business(main_row, fee_row)
                    if not shared_fee:
                        continue
                    target = int(main_row["amount_decimal"]) + fee_amount
                    for other_idx in right_by_amount.get(target, ()):
                        other_row = right.loc[other_idx]
                        if abs((pd.Timestamp(other_row["date"]) - pd.Timestamp(main_row["date"])).days) > self.config.tolerance_days:
                            continue
                        shared_other, other_basis = rows_share_business(main_row, other_row)
                        if not shared_other:
                            continue
                        bank_idxs = [main_idx, fee_idx] if left_name == "bank" else [int(other_idx)]
                        journal_idxs = [main_idx, fee_idx] if left_name == "journal" else [int(other_idx)]
                        main_positive = int(main_row["amount_decimal"]) > 0
                        if left_name == "journal" and main_positive:
                            formula = "银行实收＝账面应收－手续费"
                        elif left_name == "bank" and main_positive:
                            formula = "账面实收＝银行入账总额－手续费"
                        elif left_name == "journal":
                            formula = "银行实付＝账面应付＋手续费"
                        else:
                            formula = "账面实付＝银行扣款总额＋手续费"
                        candidate = self._add_candidate(
                            bank_idxs,
                            journal_idxs,
                            "fee_net",
                            "手续费净额",
                            resolves_full_group=True,
                            complete_business_id=True,
                            complete_business_group=True,
                            represents_full_observed_group=True,
                            business_strength=3,
                            relationship_formula=formula,
                            fee_amount_li=abs(fee_amount),
                            formula_difference_li=0,
                            business_basis=f"{fee_basis}；{other_basis}；费用文字明确",
                        )
                        if candidate is not None:
                            candidate.metrics = replace(candidate.metrics, total_diff_li=0)
                            self._score_existing_candidate(candidate)

    def _reset_candidate_search_stage(self, stage: str) -> Dict[str, int]:
        """重置一个候选阶段的可披露计数。"""
        stats = {
            "examined": 0,
            "retained": 0,
            "truncated_source_rows": 0,
        }
        self.run_parameters.setdefault("candidate_search", {})[stage] = stats
        self._candidate_search_truncated_sources[stage] = set()
        return stats

    def _record_candidate_pool(
        self,
        stage: str,
        source_key: Any,
        examined: int,
        retained: int,
        *,
        truncated: Optional[bool] = None,
    ) -> None:
        """记录过滤后的候选规模和实际保留规模，截断来源按来源键去重。"""
        stages = self.run_parameters.setdefault("candidate_search", {})
        stats = stages.setdefault(
            stage,
            {"examined": 0, "retained": 0, "truncated_source_rows": 0},
        )
        examined = max(0, int(examined))
        retained = max(0, min(int(retained), examined))
        stats["examined"] += examined
        stats["retained"] += retained
        is_truncated = examined > retained if truncated is None else bool(truncated)
        if not is_truncated:
            return
        truncated = self._candidate_search_truncated_sources.setdefault(
            stage,
            set(),
        )
        if source_key not in truncated:
            truncated.add(source_key)
            stats["truncated_source_rows"] = len(truncated)

    @staticmethod
    def _profile_pool_keys(profile: Dict[str, Any]) -> Tuple[Tuple[Any, ...], ...]:
        """提取可索引的强文字键，供截断前稳定选择候选。"""
        keys: List[Tuple[Any, ...]] = []
        keys.extend(("业务编号", *identifier) for identifier in sorted(profile.get("ids", ())))
        keys.extend(
            ("交易流水", *identifier)
            for identifier in sorted(profile.get("transaction_ids", ()))
        )
        for category, values in sorted(profile.get("parties", {}).items()):
            keys.extend(("对方", category, value) for value in sorted(values))
        summary = str(profile.get("summary", "")).strip()
        if summary:
            keys.append(("摘要", summary))
        batch_category = str(profile.get("batch_category", "")).strip()
        if batch_category:
            keys.append(("批量用途", batch_category))
        return tuple(keys)

    @staticmethod
    def _has_strong_cross_side_evidence(evidence: Dict[str, Any]) -> bool:
        """共同编号、共同流水或明确业务上下文才足以闭合跨双方范围。"""
        return bool(
            evidence.get("shared_business_id")
            or evidence.get("shared_transaction_id")
            or (
                evidence.get("same_party_group")
                and evidence.get("specific_summary_group")
            )
            or evidence.get("salary_group")
            or evidence.get("batch_category")
        )

    def _stable_profile_pool(
        self,
        source_profile: Dict[str, Any],
        target_indexes: List[int],
        target_type: str,
        limit: int,
        by_key: Optional[Dict[Tuple[Any, ...], List[int]]] = None,
    ) -> List[int]:
        """按强文字键优先、业务内容哈希兜底选出固定规模的评分池。"""
        limit = max(0, int(limit))
        if limit == 0 or not target_indexes:
            return []
        if by_key is None:
            ordered = sorted(
                (int(index) for index in target_indexes),
                key=lambda index: self._stable_row_order_key(target_type, index),
            )
            by_key = self._build_profile_pool_index(ordered, target_type)
        else:
            # 调用方已为同一金额日期组构建一次稳定顺序和索引，避免每个来源重复排序。
            ordered = [int(index) for index in target_indexes]

        source_keys = self._profile_pool_keys(source_profile)
        scores: Dict[int, int] = {}

        def shared_indexes(kind: str) -> Set[int]:
            result: Set[int] = set()
            for key in source_keys:
                if key[0] == kind:
                    result.update(int(index) for index in by_key.get(key, ()))
            return result

        business_ids = shared_indexes("业务编号")
        transaction_ids = shared_indexes("交易流水")
        summaries = shared_indexes("摘要")
        parties = shared_indexes("对方")
        batch_uses = shared_indexes("批量用途")
        for index in business_ids | transaction_ids:
            scores[index] = max(scores.get(index, 0), 50)
        for index in summaries & parties:
            scores[index] = max(scores.get(index, 0), 40)
        for index in summaries | batch_uses:
            scores[index] = max(scores.get(index, 0), 30)
        for index in parties:
            scores[index] = max(scores.get(index, 0), 20)

        stable_position = {index: position for position, index in enumerate(ordered)}
        ranked = sorted(
            ordered,
            key=lambda index: (-scores.get(index, 0), stable_position[index]),
        )
        return ranked[:limit]

    def _build_profile_pool_index(
        self,
        target_indexes: List[int],
        target_type: str,
    ) -> Dict[Tuple[Any, ...], List[int]]:
        """一次构建文字键索引，供同一金额/日期组内多个来源复用。"""
        by_key: Dict[Tuple[Any, ...], List[int]] = {}
        for index in sorted(
            (int(value) for value in target_indexes),
            key=lambda item: self._stable_row_order_key(target_type, item),
        ):
            profile = self._business_rows[target_type][index]
            for key in self._profile_pool_keys(profile):
                by_key.setdefault(key, []).append(index)
        return by_key

    def _stable_row_order_key(self, side: str, index: int) -> Tuple[str, int]:
        """截断候选时按业务内容排序；完全重复行才用内部索引打破并列。"""
        frame = self.bank if side == "bank" else self.journal
        index = int(index)
        cache = self._row_content_hash_cache[side]
        if index not in cache:
            cache[index] = self._row_content_hash(frame, index)
        return cache[index], index
        
    def _get_memory_usage_gb(self) -> float:
        if PSUTIL_AVAILABLE:
            try:
                return psutil.Process().memory_info().rss / (1024 ** 3)
            except (OSError, RuntimeError):
                pass
        return 0.0
    
    def _should_use_parallel(self, task_count: int) -> bool:
        if not PSUTIL_AVAILABLE:
            return True
        memory_gb = self._get_memory_usage_gb()
        if memory_gb > self.config.memory_limit_gb:
            self._log(f"⚠️ 内存使用 {memory_gb:.2f}GB 超过阈值 {self.config.memory_limit_gb}GB，降级为单线程")
            return False
        return True
    
    def _log(self, message: str) -> None:
        if self.logger:
            self.logger(message)
        else:
            try:
                print(message)
            except UnicodeEncodeError as exc:
                safe_message = str(message).encode(
                    exc.encoding,
                    errors="replace",
                ).decode(
                    exc.encoding,
                    errors="replace",
                )
                print(safe_message)

    def _update_progress(self, value: float) -> None:
        if self.progress_callback:
            self.progress_callback(value)
    
    def set_stopping(self, value: bool) -> None:
        self.stopping = value

    @staticmethod
    def _candidate_id(
        match_type: str,
        bank_idxs: Tuple[int, ...],
        journal_idxs: Tuple[int, ...],
    ) -> str:
        bank_part = ",".join(str(index) for index in sorted(bank_idxs))
        journal_part = ",".join(str(index) for index in sorted(journal_idxs))
        return f"{match_type}|B:{bank_part}|J:{journal_part}"

    @classmethod
    def _stable_value(cls, value: Any) -> Any:
        """把业务内容转成与输入行序无关的可哈希值。"""
        if value is None:
            return ""
        if isinstance(value, dict):
            return {
                str(key): cls._stable_value(item)
                for key, item in sorted(value.items(), key=lambda pair: str(pair[0]))
            }
        if isinstance(value, (list, tuple, set)):
            return [cls._stable_value(item) for item in value]
        if isinstance(value, (pd.Timestamp,)):
            return value.isoformat()
        if isinstance(value, Decimal):
            return format(value.normalize(), "f")
        if hasattr(value, "item"):
            try:
                return cls._stable_value(value.item())
            except (TypeError, ValueError):
                pass
        try:
            if pd.isna(value):
                return ""
        except (TypeError, ValueError):
            pass
        if isinstance(value, float):
            return format(Decimal(str(value)).normalize(), "f")
        if isinstance(value, str):
            return " ".join(unicodedata.normalize("NFKC", value).split()).casefold()
        return str(value)

    @classmethod
    def _row_content_hash(cls, frame: pd.DataFrame, index: int) -> str:
        row = frame.loc[index]
        content = {
            "date": cls._stable_value(row.get("date")),
            "amount": cls._stable_value(row.get("amount_decimal", row.get("amount"))),
            "summary": normalize_summary(str(row.get("summary", ""))),
            "voucher": cls._stable_value(row.get("voucher_no", "")),
            "voucher_word": cls._stable_value(row.get("voucher_word", "")),
            "auxiliary_text": cls._stable_value(row.get("aux_text_fields", {})),
            "amount_evidence": cls._stable_value(row.get("amount_evidence", {})),
            "voucher_evidence": cls._stable_value(row.get("voucher_evidence", {})),
        }
        payload = json.dumps(content, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()

    def _stable_candidate_identity(
        self,
        match_type: str,
        bank_idxs: Tuple[int, ...],
        journal_idxs: Tuple[int, ...],
    ) -> Tuple[str, str, str]:
        """返回候选公开编号、候选稳定键和不含候选类型的组成键。"""
        composition = {
            "bank": sorted(
                self._stable_row_order_key("bank", index)[0]
                for index in bank_idxs
            ),
            "journal": sorted(
                self._stable_row_order_key("journal", index)[0]
                for index in journal_idxs
            ),
        }
        payload = json.dumps(composition, sort_keys=True, separators=(",", ":"))
        composition_key = hashlib.sha256(payload.encode("utf-8")).hexdigest()
        stable_key = hashlib.sha256(
            f"{match_type}|{composition_key}".encode("utf-8")
        ).hexdigest()
        base_id = f"C-{stable_key[:16].upper()}"
        occurrence = self._stable_id_counts.get(base_id, 0) + 1
        self._stable_id_counts[base_id] = occurrence
        candidate_id = base_id if occurrence == 1 else f"{base_id}-{occurrence:03d}"
        return candidate_id, stable_key, composition_key

    @staticmethod
    def _row_text_fields(frame: pd.DataFrame, index: int) -> Dict[str, object]:
        if "aux_text_fields" in frame.columns:
            fields = frame.at[index, "aux_text_fields"]
            if isinstance(fields, dict):
                return dict(fields)
        summary = frame.at[index, "summary"] if "summary" in frame.columns else ""
        return {"摘要": summary}

    def _merged_text_fields(
        self,
        frame: pd.DataFrame,
        indices: Tuple[int, ...],
    ) -> Dict[str, str]:
        return merge_labeled_fields(
            [self._row_text_fields(frame, index) for index in indices]
        )

    def _score_existing_candidate(self, candidate: MatchCandidate) -> None:
        bank_fields = self._merged_text_fields(self.bank, candidate.bank_idxs)
        journal_fields = self._merged_text_fields(self.journal, candidate.journal_idxs)
        if len(candidate.bank_idxs) > 1 or len(candidate.journal_idxs) > 1:
            # 交易流水号、回单号用于核验一对一交易。批次内每笔号码本来就应不同，
            # 合并后比较两侧号码集合会把完整工资批次误判为关键字段冲突。
            bank_fields = {
                label: value for label, value in bank_fields.items()
                if _critical_field_category(str(label)) != "交易流水号"
            }
            journal_fields = {
                label: value for label, value in journal_fields.items()
                if _critical_field_category(str(label)) != "交易流水号"
            }
        scores = score_candidate(
            candidate,
            self.config,
            bank_fields,
            journal_fields,
        )
        if (
            candidate.metrics.total_diff_li == 0
            and not candidate.evidence.get("resolves_full_group")
            and int(candidate.evidence.get("business_strength", 0)) < 2
        ):
            # 仅金额、日期和相同摘要不能得到满分；结构分保留关系形态信息，
            # 同时明确反映缺少共同编号、交易流水或对方证据。
            candidate.scores = replace(
                scores,
                structure=min(scores.structure, 9),
            )
            candidate.evidence["score_limited_by_business_evidence"] = True

    def _add_candidate(
        self,
        bank_idxs: List[int] | Tuple[int, ...],
        journal_idxs: List[int] | Tuple[int, ...],
        match_type: str,
        match_stage: str,
        **evidence: Any,
    ) -> Optional[MatchCandidate]:
        """只登记候选，不提前占用银行流水或日记账记录。"""
        precomputed_business_evidence = evidence.pop(
            "_precomputed_business_evidence",
            None,
        )
        bank_tuple = tuple(sorted(int(index) for index in bank_idxs))
        journal_tuple = tuple(sorted(int(index) for index in journal_idxs))
        if not bank_tuple or not journal_tuple:
            return None
        index_key = self._candidate_id(match_type, bank_tuple, journal_tuple)
        if index_key in self._candidate_index_keys:
            return None
        bank_amounts = self.bank.loc[list(bank_tuple), "amount_decimal"].tolist()
        journal_amounts = self.journal.loc[list(journal_tuple), "amount_decimal"].tolist()
        if not self.config.allow_zero_match and (
            all(int(value) == 0 for value in bank_amounts)
            or all(int(value) == 0 for value in journal_amounts)
        ):
            return None
        candidate_id, stable_key, composition_key = self._stable_candidate_identity(
            match_type,
            bank_tuple,
            journal_tuple,
        )

        all_dates = [
            pd.Timestamp(value)
            for value in (
                self.bank.loc[list(bank_tuple), "date"].tolist()
                + self.journal.loc[list(journal_tuple), "date"].tolist()
            )
        ]
        date_span_days = (
            int((max(all_dates) - min(all_dates)).days) if all_dates else 0
        )
        for side, selected_indices in (
            ("bank", frozenset(bank_tuple)),
            ("journal", frozenset(journal_tuple)),
        ):
            for atomic_group in self._atomic_groups.get(side, []):
                overlap = selected_indices & atomic_group
                if overlap and overlap != atomic_group:
                    evidence["atomic_group_subset"] = True
                    evidence["selection_ineligible"] = True
                    evidence.setdefault(
                        "ineligible_reason",
                        "候选只包含完整凭证的一部分，禁止拆分凭证凑数",
                    )
                    break

        candidate = MatchCandidate(
            candidate_id=candidate_id,
            bank_idxs=bank_tuple,
            journal_idxs=journal_tuple,
            match_type=match_type,
            match_stage=match_stage,
            metrics=build_group_metrics(bank_amounts, journal_amounts),
            bank_dates=tuple(
                pd.Timestamp(value)
                for value in self.bank.loc[list(bank_tuple), "date"].tolist()
            ),
            journal_dates=tuple(
                pd.Timestamp(value)
                for value in self.journal.loc[list(journal_tuple), "date"].tolist()
            ),
            date_span_days=date_span_days,
            is_cross_month_many_to_many=bool(
                evidence.pop("is_cross_month_many_to_many", False)
            ),
            is_ambiguous=bool(evidence.pop("is_ambiguous", False)),
            rule_matched=bool(evidence.pop("is_rule_matched", False)),
            stable_key=stable_key,
            composition_key=composition_key,
            evidence=dict(evidence),
        )
        candidate.evidence.update(
            precomputed_business_evidence
            if precomputed_business_evidence is not None
            else business_evidence(
                [self._business_rows["bank"][index] for index in bank_tuple],
                [self._business_rows["journal"][index] for index in journal_tuple],
            )
        )
        if self.overall_scope_limited:
            candidate.evidence["overall_scope_limited"] = True
            candidate.evidence["overall_control_reasons"] = list(
                self.overall_control.reasons
            )
        self._score_existing_candidate(candidate)
        self.candidates.append(candidate)
        self._candidate_ids.add(candidate_id)
        self._candidate_index_keys.add(index_key)
        return candidate

    @staticmethod
    def _candidate_component_positions(
        candidates: List[MatchCandidate],
    ) -> List[List[int]]:
        """用并查集按共享源行形成竞争分量，避免逐候选展开整张邻接表。"""
        parent = list(range(len(candidates)))
        rank = [0] * len(candidates)

        def find(position: int) -> int:
            while parent[position] != position:
                parent[position] = parent[parent[position]]
                position = parent[position]
            return position

        def union(left: int, right: int) -> None:
            left_root = find(left)
            right_root = find(right)
            if left_root == right_root:
                return
            if rank[left_root] < rank[right_root]:
                left_root, right_root = right_root, left_root
            parent[right_root] = left_root
            if rank[left_root] == rank[right_root]:
                rank[left_root] += 1

        bank_owner: Dict[int, int] = {}
        journal_owner: Dict[int, int] = {}
        for position, candidate in enumerate(candidates):
            for index in candidate.bank_idxs:
                owner = bank_owner.setdefault(int(index), position)
                union(position, owner)
            for index in candidate.journal_idxs:
                owner = journal_owner.setdefault(int(index), position)
                union(position, owner)

        components: Dict[int, List[int]] = {}
        for position in range(len(candidates)):
            components.setdefault(find(position), []).append(position)
        result = list(components.values())
        result.sort(
            key=lambda component: min(
                candidates[position].candidate_id for position in component
            )
        )
        return result

    def _refresh_candidate_ambiguity(self) -> None:
        """候选全部生成后，以固定上限保存稳定的直接竞争关系样本。"""
        candidates = [
            candidate for candidate in self.candidates
            if not candidate.evidence.get("selection_ineligible")
        ]
        if not candidates:
            return
        conflicts = [has_business_conflict(candidate) for candidate in candidates]
        priorities = [relationship_priority(candidate) for candidate in candidates]

        def composition(position: int) -> Tuple[Any, ...]:
            candidate = candidates[position]
            return candidate.bank_idxs, candidate.journal_idxs

        representatives: Dict[Tuple[Any, ...], int] = {}
        for position, candidate in enumerate(candidates):
            if conflicts[position]:
                continue
            key = composition(position)
            current = representatives.get(key)
            if current is None or (
                priorities[position] > priorities[current]
                or (
                    priorities[position] == priorities[current]
                    and candidate.candidate_id < candidates[current].candidate_id
                )
            ):
                representatives[key] = position

        def priority_order(position: int) -> Tuple[Any, ...]:
            return (
                *( -int(value) for value in priorities[position]),
                candidates[position].candidate_id,
            )

        bank_members: Dict[int, List[int]] = {}
        journal_members: Dict[int, List[int]] = {}
        for position in representatives.values():
            candidate = candidates[position]
            for index in candidate.bank_idxs:
                bank_members.setdefault(int(index), []).append(position)
            for index in candidate.journal_idxs:
                journal_members.setdefault(int(index), []).append(position)
        for members in (*bank_members.values(), *journal_members.values()):
            members.sort(key=priority_order)

        representative_positions = sorted(representatives.values())
        representative_candidates = [
            candidates[position] for position in representative_positions
        ]
        components = [
            [representative_positions[local_position] for local_position in component]
            for component in self._candidate_component_positions(
                representative_candidates
            )
        ]
        component_by_position: Dict[int, int] = {}
        component_bank_rows: Dict[int, Set[int]] = {}
        component_journal_rows: Dict[int, Set[int]] = {}
        component_compositions: Dict[int, Set[Tuple[Any, ...]]] = {}
        for number, positions in enumerate(components):
            bank_rows: Set[int] = set()
            journal_rows: Set[int] = set()
            composition_keys: Set[Tuple[Any, ...]] = set()
            for position in positions:
                component_by_position[position] = number
                bank_rows.update(candidates[position].bank_idxs)
                journal_rows.update(candidates[position].journal_idxs)
                if not conflicts[position]:
                    composition_keys.add(composition(position))
            component_bank_rows[number] = bank_rows
            component_journal_rows[number] = journal_rows
            component_compositions[number] = composition_keys

        stored_references = 0
        for position, candidate in enumerate(candidates):
            representative_position = representatives.get(composition(position))
            if conflicts[position] or representative_position is None:
                candidate.evidence["alternative_candidate_ids"] = []
                candidate.evidence["alternative_candidate_keys"] = []
                candidate.evidence["alternative_candidate_count"] = 0
                candidate.evidence["alternative_candidates_truncated"] = False
                candidate.is_ambiguous = False
                self._score_existing_candidate(candidate)
                continue
            own_priority = priorities[position]
            weak_business_evidence = bool(
                candidate.evidence.get("business_strength", 0) <= 1
            )
            candidate_bank = set(candidate.bank_idxs)
            candidate_journal = set(candidate.journal_idxs)
            full_group = bool(
                candidate.evidence.get("resolves_full_group")
                and not conflicts[position]
                and (
                    candidate.metrics.total_diff_li == 0
                    or candidate.evidence.get("complete_business_id")
                    or candidate.evidence.get("complete_business_group")
                )
            )
            component_number = component_by_position[representative_position]
            component_count = len(component_compositions[component_number])
            if composition(position) in component_compositions[component_number]:
                component_count -= 1

            found: Dict[str, Tuple[Tuple[Any, ...], str]] = {}
            component_is_internal = bool(
                full_group
                and component_bank_rows[component_number] <= candidate_bank
                and component_journal_rows[component_number] <= candidate_journal
            )
            if not component_is_internal:
                member_lists = [
                    bank_members.get(int(index), [])
                    for index in sorted(
                        candidate.bank_idxs,
                        key=lambda item: self._stable_row_order_key("bank", int(item)),
                    )
                ] + [
                    journal_members.get(int(index), [])
                    for index in sorted(
                        candidate.journal_idxs,
                        key=lambda item: self._stable_row_order_key("journal", int(item)),
                    )
                ]
                for members in member_lists:
                    accepted_in_member = 0
                    for other_position in members:
                        other = candidates[other_position]
                        other_priority = priorities[other_position]
                        compared_other = (
                            other_priority[:-1]
                            if weak_business_evidence
                            else other_priority
                        )
                        compared_own = (
                            own_priority[:-1]
                            if weak_business_evidence
                            else own_priority
                        )
                        if compared_other < compared_own:
                            break
                        if composition(other_position) == composition(position):
                            continue
                        if (
                            full_group
                            and set(other.bank_idxs) <= candidate_bank
                            and set(other.journal_idxs) <= candidate_journal
                        ):
                            continue
                        found[other.candidate_id] = (
                            priority_order(other_position),
                            other.stable_key,
                        )
                        accepted_in_member += 1
                        if accepted_in_member >= self.MAX_ALTERNATIVE_REFERENCES:
                            break

            selected = sorted(
                (
                    (order, candidate_id, stable_key)
                    for candidate_id, (order, stable_key) in found.items()
                ),
                key=lambda item: (item[0], item[1]),
            )[: self.MAX_ALTERNATIVE_REFERENCES]
            alternative_ids = [item[1] for item in selected]
            alternative_keys = [item[2] for item in selected]
            candidate.evidence["alternative_candidate_ids"] = alternative_ids
            candidate.evidence["alternative_candidate_keys"] = alternative_keys
            candidate.evidence["alternative_candidate_count"] = (
                component_count if alternative_ids else 0
            )
            candidate.evidence["alternative_candidates_truncated"] = bool(
                alternative_ids and component_count > len(alternative_ids)
            )
            candidate.is_ambiguous = bool(alternative_ids)
            stored_references += len(alternative_ids)
            self._score_existing_candidate(candidate)
        self.run_parameters["candidate_ambiguity"] = {
            "candidate_count": len(candidates),
            "component_count": len(components),
            "reference_limit_per_candidate": self.MAX_ALTERNATIVE_REFERENCES,
            "stored_reference_count": stored_references,
        }

    def _group_ambiguous_candidates(
        self,
        candidates: Optional[List[MatchCandidate]] = None,
    ) -> List[List[MatchCandidate]]:
        """把共享任一银行或日记账记录的候选归入同一竞争组。"""
        candidates = (
            [
                candidate for candidate in self.candidates
                if not candidate.evidence.get("selection_ineligible")
            ]
            if candidates is None
            else candidates
        )
        return [
            [
                candidates[position]
                for position in sorted(
                    component,
                    key=lambda item: candidates[item].candidate_id,
                )
            ]
            for component in self._candidate_component_positions(candidates)
        ]

    def _add_closed_candidate_groups(self) -> None:
        """把多解但整体借贷相等的竞争组归并为可自动确认的整组关系。"""
        exact_candidates = [
            candidate
            for candidate in self.candidates
            if candidate.metrics.total_diff_li == 0
            and not candidate.evidence.get("closed_group_fallback")
            and not candidate.evidence.get("selection_ineligible")
            and not has_business_conflict(candidate)
        ]
        for group in self._group_ambiguous_candidates(exact_candidates):
            if len(group) < 2:
                continue
            distinct_compositions = {
                (candidate.bank_idxs, candidate.journal_idxs)
                for candidate in group
            }
            if len(distinct_compositions) < 2:
                # 同一组成只是由多个匹配阶段重复生成，不构成真正的多解封闭组。
                continue
            if any(
                candidate.evidence.get("resolves_full_group")
                for candidate in group
            ):
                continue
            bank_idxs = sorted(
                {index for candidate in group for index in candidate.bank_idxs}
            )
            journal_idxs = sorted(
                {index for candidate in group for index in candidate.journal_idxs}
            )
            group_evidence = business_evidence(
                [self._business_rows["bank"][index] for index in bank_idxs],
                [self._business_rows["journal"][index] for index in journal_idxs],
            )
            same_amount_multiset = (
                len(bank_idxs) == len(journal_idxs) > 1
                and sorted(
                    int(value)
                    for value in self.bank.loc[bank_idxs, "amount_decimal"]
                ) == sorted(
                    int(value)
                    for value in self.journal.loc[journal_idxs, "amount_decimal"]
                )
            )
            one_side_is_single = len(bank_idxs) == 1 or len(journal_idxs) == 1
            has_boundary = bool(
                not group_evidence["business_conflicts"]
                and (
                    group_evidence["shared_business_id"]
                    or (
                        one_side_is_single
                        and (
                            group_evidence["salary_group"]
                            or group_evidence["batch_category"]
                            or group_evidence["specific_summary_group"]
                            or group_evidence["business_strength"] >= 3
                        )
                    )
                    or (
                        same_amount_multiset
                        and group_evidence.get("same_party_group", False)
                        and group_evidence.get("specific_summary_group", False)
                        and not group_evidence["batch_category"]
                    )
                )
            )
            if not has_boundary:
                continue
            if not self._total_structure_matches(
                self.bank.loc[bank_idxs, "amount_decimal"].tolist(),
                self.journal.loc[journal_idxs, "amount_decimal"].tolist(),
            ):
                continue
            self._add_candidate(
                bank_idxs,
                journal_idxs,
                "closed_candidate_group",
                "整组勾稽",
                resolves_full_group=True,
                complete_business_group=True,
                represents_full_observed_group=True,
                closed_group_fallback=True,
                alternative_count=len(group),
            )

    @staticmethod
    def _candidate_sort_key(candidate: MatchCandidate) -> Tuple[Any, ...]:
        return candidate_sort_key(candidate)

    def _llm_candidate_limit(self) -> int:
        configured = getattr(self.llm_assistant, "candidate_limit", None)
        assistant_config = getattr(self.llm_assistant, "config", None)
        if assistant_config is not None:
            configured = getattr(
                assistant_config,
                "candidate_limit",
                configured,
            )
        try:
            return max(1, int(configured if configured is not None else 5))
        except (TypeError, ValueError):
            return 5

    def _semantic_candidate(
        self,
        candidate: MatchCandidate,
    ) -> SemanticCandidate:
        bank_fields = self._merged_text_fields(
            self.bank,
            candidate.bank_idxs,
        )
        journal_fields = self._merged_text_fields(
            self.journal,
            candidate.journal_idxs,
        )
        bank_date = (
            min(candidate.bank_dates).date().isoformat()
            if candidate.bank_dates
            else ""
        )
        journal_date = (
            min(candidate.journal_dates).date().isoformat()
            if candidate.journal_dates
            else ""
        )
        return SemanticCandidate(
            candidate_id=candidate.candidate_id,
            bank_date=bank_date,
            journal_date=journal_date,
            bank_amount=float(
                PrecisionEngine.from_integer_li(
                    candidate.metrics.bank_gross_li
                )
            ),
            journal_amount=float(
                PrecisionEngine.from_integer_li(
                    candidate.metrics.journal_gross_li
                )
            ),
            bank_fields=bank_fields,
            journal_fields=journal_fields,
            local_signals=build_sensitive_field_signals(
                bank_fields,
                journal_fields,
            ),
        )

    def _apply_llm_assistance(self) -> None:
        """只增强含糊候选的文字分，硬性金额和状态规则保持不变。"""
        if self.llm_assistant is None:
            return
        assistant_config = getattr(self.llm_assistant, "config", None)
        if assistant_config is not None and not getattr(
            assistant_config,
            "enabled",
            True,
        ):
            return

        performance_li = PrecisionEngine.to_integer_li(
            self.config.performance_materiality
        )
        request_sequence = len(self.llm_records)
        for group in self._group_ambiguous_candidates():
            eligible = [
                candidate
                for candidate in group
                if candidate.metrics.group_amount_li <= performance_li
            ]
            if not eligible:
                continue
            if (
                len(eligible) == 1
                and eligible[0].match_type == "exact_1to1"
                and eligible[0].metrics.total_diff_li == 0
            ):
                continue
            needs_llm = (
                len(eligible) > 1
                or any(
                    abs(
                        candidate.scores.total
                        - self.config.auto_confirm_score
                    )
                    <= 10
                    for candidate in eligible
                )
                or any(
                    candidate.text_evidence is None
                    or candidate.text_evidence.score < 10
                    for candidate in eligible
                )
            )
            if not needs_llm:
                continue

            limited = sorted(
                eligible,
                key=self._candidate_sort_key,
            )[: self._llm_candidate_limit()]
            request_sequence += 1
            semantic_request = CandidateSemanticRequest(
                request_id=f"LLM{request_sequence:06d}",
                candidates=tuple(
                    self._semantic_candidate(candidate)
                    for candidate in limited
                ),
            )
            for candidate in limited:
                candidate.evidence.setdefault(
                    "pre_llm_total_score",
                    candidate.scores.total,
                )
            try:
                decision = self.llm_assistant.evaluate_candidates(
                    semantic_request
                )
            except Exception as exc:
                decision = LLMDecisionRecord(
                    request_id=semantic_request.request_id,
                    candidate_ids=tuple(
                        candidate.candidate_id
                        for candidate in limited
                    ),
                    fallback_used=True,
                    error=f"大模型辅助异常，已使用本地文字评分：{exc}",
                )

            if not decision.candidate_ids:
                decision.candidate_ids = tuple(
                    candidate.candidate_id for candidate in limited
                )
            if not decision.sent_fields:
                allowed_fields = set(ONLINE_ALLOWED_FIELDS)
                if (
                    assistant_config is not None
                    and getattr(assistant_config, "mode", "") == "local"
                ):
                    allowed_fields = set(
                        getattr(assistant_config, "local_fields", ())
                    )
                submitted_fields = {"日期", "金额"}
                for semantic_candidate in semantic_request.candidates:
                    submitted_fields.update(
                        set(semantic_candidate.bank_fields) & allowed_fields
                    )
                    submitted_fields.update(
                        set(semantic_candidate.journal_fields)
                        & allowed_fields
                    )
                    if semantic_candidate.local_signals:
                        submitted_fields.add(
                            "本机敏感字段一致性信号"
                        )
                decision.sent_fields = tuple(sorted(submitted_fields))
            if not decision.protocol and assistant_config is not None:
                decision.protocol = str(
                    getattr(assistant_config, "protocol", "")
                )

            valid_ids = {candidate.candidate_id for candidate in limited}
            if (
                not decision.fallback_used
                and decision.selected_candidate_id not in valid_ids
            ):
                decision.fallback_used = True
                decision.error = "模型返回了候选集以外编号，已使用本地文字评分"
                decision.selected_candidate_id = ""
            self.llm_records.append(decision)
            if decision.fallback_used:
                continue

            selected = next(
                candidate
                for candidate in limited
                if candidate.candidate_id
                == decision.selected_candidate_id
            )
            selected.llm_decision = decision
            score_candidate(
                selected,
                self.config,
                self._merged_text_fields(
                    self.bank,
                    selected.bank_idxs,
                ),
                self._merged_text_fields(
                    self.journal,
                    selected.journal_idxs,
                ),
                llm_semantic_score=decision.semantic_score,
            )

    @staticmethod
    def _legacy_confidence(score: int) -> str:
        if score >= 85:
            return "高"
        if score >= 70:
            return "中"
        return "低"

    def _commit_selected_candidates(self) -> None:
        """统一完成排序、占用、稳定编号和旧字段回填。"""
        self._add_closed_candidate_groups()
        self._refresh_candidate_ambiguity()
        self._apply_llm_assistance()
        selection_diagnostics: Dict[str, Any] = {}
        self.selected_candidates = select_non_conflicting_candidates(
            self.candidates,
            diagnostics=selection_diagnostics,
            should_stop=lambda: self.stopping,
        )
        self.run_parameters["selection_optimization"] = selection_diagnostics
        final_id_counts: Dict[str, int] = {}
        for candidate in self.selected_candidates:
            base_final_id = f"M-{candidate.composition_key[:16].upper()}"
            occurrence = final_id_counts.get(base_final_id, 0) + 1
            final_id_counts[base_final_id] = occurrence
            candidate.final_match_id = (
                base_final_id
                if occurrence == 1
                else f"{base_final_id}-{occurrence:03d}"
            )
            status, risk, reason = route_candidate(candidate, self.config)
            candidate.processing_status = status
            candidate.risk_level = risk
            candidate.processing_reason = reason
            confidence = self._legacy_confidence(candidate.scores.total)
            bank_idxs = list(candidate.bank_idxs)
            journal_idxs = list(candidate.journal_idxs)

            for frame, indices in (
                (self.bank, bank_idxs),
                (self.journal, journal_idxs),
            ):
                frame.loc[indices, "matched"] = True
                frame.loc[indices, "match_id"] = candidate.final_match_id
                frame.loc[indices, "match_type"] = candidate.match_type
                frame.loc[indices, "confidence"] = confidence
                frame.loc[indices, "confidence_score"] = candidate.scores.total
                frame.loc[indices, "processing_status"] = status.value
                frame.loc[indices, "risk_level"] = risk.value

            self.matches.append(
                {
                    "id": candidate.final_match_id,
                    "type": candidate.match_type,
                    "confidence": confidence,
                    "confidence_score": candidate.scores.total,
                    "processing_status": status.value,
                    "risk_level": risk.value,
                    "processing_reason": reason,
                    "bank_idxs": bank_idxs,
                    "journal_idxs": journal_idxs,
                    "match_stage": candidate.match_stage,
                    "amount_diff": PrecisionEngine.from_integer_li(
                        candidate.metrics.total_diff_li
                    ),
                    "date_diff_days": candidate.date_span_days,
                    "summary_similarity": (
                        candidate.text_evidence.local_score / 100
                        if candidate.text_evidence
                        else 0.0
                    ),
                    "is_rule_matched": candidate.rule_matched,
                    "is_tolerance_matched": (
                        candidate.match_type == "tolerance_date"
                    ),
                    "combo_count": len(bank_idxs) + len(journal_idxs),
                    "is_aggregation_matched": (
                        len(bank_idxs) > 1 or len(journal_idxs) > 1
                    ),
                    "score_breakdown": candidate.scores,
                    "text_evidence": candidate.text_evidence,
                }
            )
        self.difference_pools = apply_monthly_difference_pools(
            self.selected_candidates,
            self.config,
        )
        match_by_candidate = {
            candidate.candidate_id: match
            for candidate, match in zip(
                self.selected_candidates,
                self.matches,
            )
        }
        for candidate in self.selected_candidates:
            if not candidate.evidence.get(
                "included_in_risk_pool",
                False,
            ):
                continue
            for frame, indices in (
                (self.bank, list(candidate.bank_idxs)),
                (self.journal, list(candidate.journal_idxs)),
            ):
                frame.loc[
                    indices,
                    "risk_level",
                ] = candidate.risk_level.value
            match = match_by_candidate.get(candidate.candidate_id)
            if match is not None:
                match["risk_level"] = candidate.risk_level.value
                match["processing_reason"] = candidate.processing_reason

    def _commit_if_standalone(self) -> None:
        """兼容直接调用单个匹配阶段的旧用法。"""
        if not self._collecting_candidates:
            self._commit_selected_candidates()

    @staticmethod
    def _confidence_sort_key(confidence: str) -> int:
        order = {'高': 0, '中': 1, '低': 2}
        return order.get(confidence, 99)

    def _total_structure_matches(self, bank_amounts: List[Any], journal_amounts: List[Any]) -> bool:
        """收入和支出分别一致即可，笔数与金额分布只作为评分证据。"""
        metrics = build_group_metrics(bank_amounts, journal_amounts)
        return metrics.income_diff_li == 0 and metrics.expense_diff_li == 0

    def _check_negative_rules(self, summary: str) -> bool:
        negative_rules = self.MATCHING_RULES.get('negative_rules', {})
        for rule_name, keywords in negative_rules.items():
            pat = "|".join(map(re.escape, keywords))
            if re.search(pat, summary, re.IGNORECASE):
                return True
        return False

    def _rule_search_text(self, frame: pd.DataFrame, index: int) -> str:
        fields = self._row_text_fields(frame, index)
        return " ".join(str(value) for value in fields.values() if pd.notna(value))

    def _generate_rule_candidates(
        self,
        rules: Dict[str, List[str]],
        type_prefix: str,
        allowed_offsets: List[int],
    ) -> None:
        max_offset = max(allowed_offsets, default=0)
        for rule_name, keywords in rules.items():
            pattern = re.compile("|".join(map(re.escape, keywords)), re.IGNORECASE)
            bank_indexes = []
            for index in self.bank.index:
                index = int(index)
                search_text = self._rule_search_text(self.bank, index)
                if pattern.search(search_text) and not self._check_negative_rules(search_text):
                    bank_indexes.append(index)
            journal_indexes = []
            for index in self.journal.index:
                index = int(index)
                search_text = self._rule_search_text(self.journal, index)
                if pattern.search(search_text) and not self._check_negative_rules(search_text):
                    journal_indexes.append(index)

            journal_by_amount: Dict[int, Dict[str, Any]] = {}
            for journal_index in journal_indexes:
                amount = int(self.journal.at[journal_index, "amount_decimal"])
                date = pd.Timestamp(self.journal.at[journal_index, "date"])
                bucket = journal_by_amount.setdefault(
                    amount,
                    {"entries": [], "by_key": {}},
                )
                bucket["entries"].append((date, journal_index))
                profile = self._business_rows["journal"][journal_index]
                for profile_key in self._profile_pool_keys(profile):
                    bucket["by_key"].setdefault(profile_key, []).append(
                        (date, journal_index)
                    )
            for bucket in journal_by_amount.values():
                bucket["entries"].sort(
                    key=lambda item: (
                        item[0],
                        self._stable_row_order_key("journal", item[1]),
                    )
                )
                bucket["dates"] = [item[0] for item in bucket["entries"]]
                for profile_key, entries in bucket["by_key"].items():
                    entries.sort(
                        key=lambda item: (
                            item[0],
                            self._stable_row_order_key("journal", item[1]),
                        )
                    )
                    bucket["by_key"][profile_key] = (
                        entries,
                        [item[0] for item in entries],
                    )

            for bank_index in sorted(bank_indexes):
                bank_amount = int(self.bank.at[bank_index, "amount_decimal"])
                bank_date = pd.Timestamp(self.bank.at[bank_index, "date"])
                bucket = journal_by_amount.get(bank_amount)
                if bucket is None:
                    continue
                date_min = bank_date - timedelta(days=max_offset)
                date_max = bank_date + timedelta(days=max_offset)
                start = bisect.bisect_left(bucket["dates"], date_min)
                end = bisect.bisect_right(bucket["dates"], date_max)
                examined = end - start
                if examined <= 0:
                    continue
                limit = max(0, int(self.config.max_candidates))
                selected: List[int] = []
                seen: Set[int] = set()
                source_profile = self._business_rows["bank"][bank_index]
                if limit:
                    for profile_key in self._profile_pool_keys(source_profile):
                        entries_and_dates = bucket["by_key"].get(profile_key)
                        if entries_and_dates is None:
                            continue
                        entries, dates = entries_and_dates
                        key_start = bisect.bisect_left(dates, date_min)
                        key_end = bisect.bisect_right(dates, date_max)
                        for position in range(key_start, key_end):
                            _, journal_index = entries[position]
                            if journal_index in seen:
                                continue
                            selected.append(journal_index)
                            seen.add(journal_index)
                            if len(selected) >= limit:
                                break
                        if len(selected) >= limit:
                            break
                if len(selected) < limit:
                    for position in range(start, end):
                        _, journal_index = bucket["entries"][position]
                        if journal_index in seen:
                            continue
                        selected.append(journal_index)
                        seen.add(journal_index)
                        if len(selected) >= limit:
                            break
                self._record_candidate_pool(
                    "whitelist",
                    int(bank_index),
                    examined,
                    len(selected),
                )
                possible = []
                for journal_index in selected:
                    journal_date = pd.Timestamp(self.journal.at[journal_index, "date"])
                    date_difference = abs((bank_date - journal_date).days)
                    evidence = score_text_fields(
                        self._row_text_fields(self.bank, int(bank_index)),
                        self._row_text_fields(self.journal, int(journal_index)),
                    )
                    possible.append(
                        (
                            -evidence.local_score,
                            date_difference,
                            int(journal_index),
                        )
                    )
                for _, _, journal_index in sorted(possible):
                    self._add_candidate(
                        [int(bank_index)],
                        [journal_index],
                        f"{type_prefix}_{rule_name}",
                        "白名单",
                        is_rule_matched=True,
                        rule_name=rule_name,
                        structure_bonus=(type_prefix != "弱规则"),
                    )

    def match_whitelist_rules(self) -> None:
        """白名单只缩小候选范围，不能跳过金额、方向和统一评分。"""
        self._reset_candidate_search_stage("whitelist")
        self._log("  生成强规则候选...")
        self._generate_rule_candidates(
            self.MATCHING_RULES.get("strong_rules", {}),
            "强规则",
            DATE_OFFSET_PRIORITIES,
        )
        self._log("  生成正规则候选...")
        self._generate_rule_candidates(
            self.MATCHING_RULES.get("positive_rules", {}),
            "正规则",
            DATE_OFFSET_PRIORITIES,
        )
        self._log("  生成弱规则候选...")
        self._generate_rule_candidates(
            self.MATCHING_RULES.get("weak_rules", {}),
            "弱规则",
            [0, 1],
        )

    def run(self) -> List[Dict[str, Any]]:
        candidate_steps = [
            ("退款冲销重付、重复与手续费净额", self.match_special_business_events),
            ("业务完整组匹配", self.match_business_groups),
            ("白名单规则匹配", self.match_whitelist_rules),
            ("精确匹配", self.match_exact_1to1),
            ("日期容差匹配", self.match_tolerance),
            ("批量聚合匹配", self.match_batch_aggregation),
            ("连续摘要整组匹配", self.match_continuous_summary_groups),
            ("智能组合匹配", self.match_dfs_combinations),
            ("日总额匹配", self.match_daily_total),
            ("月度总额匹配", self.match_monthly_total),
            ("跨月多对多匹配", self.match_cross_month_total),
        ]

        self._collecting_candidates = True
        try:
            step_count = len(candidate_steps)
            for step_number, (name, func) in enumerate(candidate_steps, 1):
                self._log(f"开始: {name}...")
                if self.stopping:
                    self._log("任务已停止")
                    break
                stage_started = time.perf_counter()
                candidates_before = len(self.candidates)
                func()
                _gc_cleanup(f"{name}后", self._log)
                candidates_after = len(self.candidates)
                self._update_progress(step_number / step_count * 90)
                self._log(
                    f"完成: {name}；新增候选 "
                    f"{candidates_after - candidates_before:,} 组，"
                    f"累计候选 {candidates_after:,} 组，"
                    f"耗时 {time.perf_counter() - stage_started:.1f} 秒"
                )
        finally:
            self._collecting_candidates = False

        if not self.stopping:
            selection_started = time.perf_counter()
            self._log(
                f"开始: 候选关系去重与冲突选择；"
                f"待处理候选 {len(self.candidates):,} 组"
            )
            self._commit_selected_candidates()
            covered_bank = sum(
                len(candidate.bank_idxs) for candidate in self.selected_candidates
            )
            covered_journal = sum(
                len(candidate.journal_idxs) for candidate in self.selected_candidates
            )
            self._update_progress(100)
            self._log(
                f"候选选择完成：选中关系 {len(self.selected_candidates):,} 组，"
                f"覆盖银行流水 {covered_bank:,}/{len(self.bank):,} 行，"
                f"覆盖银行存款序时账 {covered_journal:,}/{len(self.journal):,} 行，"
                f"耗时 {time.perf_counter() - selection_started:.1f} 秒"
            )

        return self.matches

    def match_business_groups(self) -> None:
        groups = {side: complete_groups(rows, self.config.dfs_date_window, side)
                  for side, rows in self._business_rows.items()}
        self._reset_candidate_search_stage("business_group")
        self._atomic_groups = {
            side: [
                frozenset(indices)
                for group_key, indices in side_groups
                if group_key[0] == "凭证" and len(indices) > 1
            ]
            for side, side_groups in groups.items()
        }
        journal_by_profile_key: Dict[
            Tuple[Any, ...], List[Tuple[int, ...]]
        ] = {}
        journal_by_profile_key_total: Dict[
            Tuple[Tuple[Any, ...], int], List[Tuple[int, ...]]
        ] = {}
        journal_group_keys: Dict[Tuple[int, ...], Any] = {}
        journal_group_totals: Dict[Tuple[int, ...], int] = {}
        journal_group_dates: Dict[
            Tuple[int, ...], Tuple[pd.Timestamp, pd.Timestamp]
        ] = {}
        journal_amount_masks: Dict[int, int] = {}
        journal_profile_masks: Dict[Tuple[Any, ...], int] = {}
        journal_profile_total_masks: Dict[
            Tuple[Tuple[Any, ...], int], int
        ] = {}

        def batch_signature(side, indices):
            rows = [self._business_rows[side][i] for i in indices]
            categories = {row.get("batch_category", "") for row in rows}
            if not rows or "" in categories or len(categories) != 1:
                return None
            periods = set().union(*(row["periods"] for row in rows))
            signs = {row["sign"] for row in rows}
            if len(signs) != 1:
                return None
            if periods:
                period_key = ("期间", tuple(sorted(periods)))
            else:
                period_key = ("日期", min(row["date"].normalize() for row in rows))
            return next(iter(categories)), period_key, next(iter(signs))

        def group_profile_keys(
            side: str,
            indices: Tuple[int, ...],
        ) -> Tuple[Tuple[Any, ...], ...]:
            keys = {
                key
                for index in indices
                for key in self._profile_pool_keys(
                    self._business_rows[side][index]
                )
            }
            return tuple(sorted(keys, key=lambda value: tuple(map(str, value))))

        def profile_key_rank(key: Tuple[Any, ...]) -> int:
            return {
                "业务编号": 0,
                "交易流水": 0,
                "对方": 1,
                "批量用途": 2,
                "摘要": 3,
            }.get(str(key[0]) if key else "", 4)

        for ordinal, (group_key, indices) in enumerate(groups["journal"]):
            journal_group_keys[indices] = group_key
            total = sum(int(self.journal.at[i, "amount_decimal"]) for i in indices)
            dates = [
                pd.Timestamp(self._business_rows["journal"][index]["date"])
                for index in indices
            ]
            mask = 1 << ordinal
            journal_group_totals[indices] = total
            journal_group_dates[indices] = (min(dates), max(dates))
            journal_amount_masks[total] = journal_amount_masks.get(total, 0) | mask
            for profile_key in group_profile_keys("journal", indices):
                journal_by_profile_key.setdefault(profile_key, []).append(indices)
                key_total = (profile_key, total)
                journal_by_profile_key_total.setdefault(
                    key_total, []
                ).append(indices)
                journal_profile_masks[profile_key] = (
                    journal_profile_masks.get(profile_key, 0) | mask
                )
                journal_profile_total_masks[key_total] = (
                    journal_profile_total_masks.get(key_total, 0) | mask
                )

        def journal_group_order(indices: Tuple[int, ...]) -> Tuple[Any, ...]:
            return (
                journal_group_dates[indices][0],
                tuple(
                    self._stable_row_order_key("journal", index)
                    for index in indices
                ),
            )

        indexed_group_buckets = (
            *journal_by_profile_key.values(),
            *journal_by_profile_key_total.values(),
        )
        for indexed_groups in indexed_group_buckets:
            indexed_groups.sort(
                key=journal_group_order
            )
        journal_group_start_dates = {
            id(indexed_groups): [
                journal_group_dates[indices][0]
                for indices in indexed_groups
            ]
            for indexed_groups in indexed_group_buckets
        }

        def bounded_journal_groups(
            indexed_groups: List[Tuple[int, ...]],
            date_min: pd.Timestamp,
            date_max: pd.Timestamp,
            center: pd.Timestamp,
            limit: int,
        ) -> List[Tuple[int, ...]]:
            if not indexed_groups or limit <= 0:
                return []
            dates = journal_group_start_dates[id(indexed_groups)]
            start = bisect.bisect_left(dates, date_min)
            end = bisect.bisect_right(dates, date_max)
            if start >= end:
                return []
            middle = bisect.bisect_left(dates, center, start, end)
            sample_limit = limit * 2
            left = max(start, middle - sample_limit // 2)
            right = min(end, left + sample_limit)
            left = max(start, right - sample_limit)
            return [
                indices
                for indices in indexed_groups[left:right]
                if journal_group_dates[indices][1] <= date_max
            ]
        for bank_group_key, bank_idxs in groups["bank"]:
            if self.stopping:
                return
            total = sum(int(self.bank.at[i, "amount_decimal"]) for i in bank_idxs)
            if not total and not self.config.allow_zero_match:
                continue
            rows = [self._business_rows["bank"][i] for i in bank_idxs]
            batch_key = batch_signature("bank", bank_idxs)
            candidate_limit = max(0, int(self.config.max_candidates))
            possible_scores: Dict[Tuple[int, ...], int] = {}
            bank_dates = [pd.Timestamp(row["date"]) for row in rows]
            allowed_journal_date_min = max(bank_dates) - timedelta(
                days=self.config.dfs_date_window
            )
            allowed_journal_date_max = min(bank_dates) + timedelta(
                days=self.config.dfs_date_window
            )
            bank_date_center = min(bank_dates) + (
                max(bank_dates) - min(bank_dates)
            ) / 2
            bank_profile_keys = sorted(
                group_profile_keys("bank", bank_idxs),
                key=lambda key: (profile_key_rank(key), tuple(map(str, key))),
            )
            examined_mask = journal_amount_masks.get(total, 0)
            for profile_key in bank_profile_keys:
                rank = profile_key_rank(profile_key)
                key_total = (profile_key, total)
                strong_key = bool(
                    profile_key
                    and str(profile_key[0])
                    in {"业务编号", "交易流水", "批量用途"}
                )
                exact_total_groups = journal_by_profile_key_total.get(
                    key_total, []
                )
                examined_mask |= journal_profile_total_masks.get(
                    key_total, 0
                )
                candidate_buckets = [exact_total_groups]
                if strong_key:
                    examined_mask |= journal_profile_masks.get(profile_key, 0)
                    candidate_buckets.append(
                        journal_by_profile_key.get(profile_key, [])
                    )
                for indexed_groups in candidate_buckets:
                    for journal_idxs in bounded_journal_groups(
                        indexed_groups,
                        allowed_journal_date_min,
                        allowed_journal_date_max,
                        bank_date_center,
                        candidate_limit,
                    ):
                        possible_scores[journal_idxs] = min(
                            rank,
                            possible_scores.get(journal_idxs, rank),
                        )
            possible = sorted(
                possible_scores,
                key=lambda journal_idxs: (
                    possible_scores[journal_idxs],
                    abs(
                        total
                        - journal_group_totals[journal_idxs]
                    ),
                    abs(
                        (
                            journal_group_dates[journal_idxs][0]
                            - bank_date_center
                        ).days
                    ),
                    tuple(
                        self._stable_row_order_key("journal", index)
                        for index in journal_idxs
                    ),
                ),
            )[:candidate_limit]
            examined_count = examined_mask.bit_count()
            evaluated_count = 0
            for journal_idxs in possible:
                journal_group_key = journal_group_keys[journal_idxs]
                others = [self._business_rows["journal"][i] for i in journal_idxs]
                evaluated_count += 1
                evidence = business_evidence(rows, others)
                if evidence["business_conflicts"] or not evidence["business_strength"]:
                    continue
                if len(bank_idxs) == len(journal_idxs) == 1 and (
                        not evidence["shared_business_id"] or
                        total == int(self.journal.at[journal_idxs[0], "amount_decimal"])):
                    continue
                if {row["sign"] for row in rows + others} not in ({1}, {-1}, {0}):
                    continue
                if (evidence["business_strength"] <= 1
                        and len(bank_idxs) == len(self.bank)
                        and len(journal_idxs) == len(self.journal)):
                    # 全表普通摘要组沿用日/月总额类型，保持现有报告分类。
                    continue
                dates = [row["date"] for row in rows + others]
                if (max(dates) - min(dates)).days > self.config.dfs_date_window:
                    continue
                journal_total = sum(
                    int(self.journal.at[i, "amount_decimal"])
                    for i in journal_idxs
                )
                journal_batch_key = batch_signature("journal", journal_idxs)
                explicit_batch_group = bool(
                    evidence.get("shared_business_category") == "批次"
                )
                same_batch_group = bool(
                    explicit_batch_group
                    or (
                        batch_key is not None
                        and batch_key == journal_batch_key
                    )
                )
                complete_voucher_group = bool(
                    journal_group_key[0] == "凭证"
                    and len(journal_idxs) > 1
                )
                strong_cross_side_evidence = (
                    self._has_strong_cross_side_evidence(evidence)
                    or (
                        complete_voucher_group
                        and evidence.get("same_party_group")
                    )
                )
                batch_boundary_uncertain = bool(
                    len(bank_idxs) > 1
                    and len(journal_idxs) > 1
                    and not evidence["shared_business_id"]
                    and not complete_voucher_group
                )
                complete_business_group = bool(
                    evidence["shared_business_id"]
                    or (
                        complete_voucher_group
                        and strong_cross_side_evidence
                    )
                    or (same_batch_group and not batch_boundary_uncertain)
                )
                if explicit_batch_group:
                    business_group_kind = "批次"
                elif evidence["shared_business_id"]:
                    business_group_kind = "编号"
                elif complete_voucher_group:
                    business_group_kind = "凭证"
                elif same_batch_group:
                    business_group_kind = batch_key[0]
                else:
                    business_group_kind = bank_group_key[0]
                candidate = self._add_candidate(
                    bank_idxs,
                    journal_idxs,
                    "business_group",
                    "业务完整组",
                    resolves_full_group=bool(
                        not batch_boundary_uncertain
                        and strong_cross_side_evidence
                    ),
                    is_rule_matched=True,
                    complete_business_id=evidence["shared_business_id"],
                    complete_business_group=complete_business_group,
                    represents_full_observed_group=True,
                    business_group_kind=business_group_kind,
                    batch_boundary_uncertain=batch_boundary_uncertain,
                    batch_difference=bool(
                        same_batch_group and total != journal_total
                    ),
                )
                if candidate is not None and same_batch_group:
                    count_difference = len(candidate.bank_idxs) - len(candidate.journal_idxs)
                    amount_gap = PrecisionEngine.from_integer_li(
                        candidate.metrics.total_diff_li
                    )
                    bank_count = len(candidate.bank_idxs)
                    journal_count = len(candidate.journal_idxs)
                    if bank_count == 1 and journal_count > 1:
                        count_hint = (
                            f"银行流水为1笔汇总，银行日记账为{journal_count}笔明细；"
                            f"两侧记录层级不同，不能按{journal_count - 1}笔差直接判断缺项"
                        )
                    elif journal_count == 1 and bank_count > 1:
                        count_hint = (
                            f"银行流水为{bank_count}笔明细，银行日记账为1笔汇总；"
                            f"两侧记录层级不同，不能按{bank_count - 1}笔差直接判断缺项"
                        )
                    elif count_difference < 0:
                        count_hint = (
                            f"银行流水较银行日记账少{abs(count_difference)}笔记录"
                            f"（银行{bank_count}笔、日记账{journal_count}笔）；"
                            "双方均为明细时也不能仅凭记录笔数认定缺项或重复"
                        )
                    elif count_difference > 0:
                        count_hint = (
                            f"银行流水较银行日记账多{count_difference}笔记录"
                            f"（银行{bank_count}笔、日记账{journal_count}笔）；"
                            "双方均为明细时也不能仅凭记录笔数认定缺项或重复"
                        )
                    else:
                        count_hint = "双方批次笔数相同"
                    review_hint = ""
                    if candidate.metrics.total_diff_li or count_difference:
                        bank_gross = candidate.metrics.bank_gross_li
                        journal_gross = candidate.metrics.journal_gross_li
                        if bank_gross < journal_gross:
                            amount_hint = f"银行流水合计较银行日记账少{amount_gap}元"
                            possible_side = "银行流水金额较少"
                        elif bank_gross > journal_gross:
                            amount_hint = f"银行流水合计较银行日记账多{amount_gap}元"
                            possible_side = "银行日记账金额较少"
                        else:
                            amount_hint = "双方批次总额相同"
                            possible_side = "未定位"
                        detailed_frame = (
                            self.bank.loc[list(candidate.bank_idxs)]
                            if bank_count > 1 and journal_count == 1
                            else (
                                self.journal.loc[list(candidate.journal_idxs)]
                                if journal_count > 1 and bank_count == 1
                                else None
                            )
                        )
                        estimate_hint = ""
                        if detailed_frame is not None and candidate.metrics.total_diff_li:
                            common_amounts = [
                                abs(int(value))
                                for value in detailed_frame["amount_decimal"].tolist()
                                if int(value) != 0
                            ]
                            if common_amounts:
                                frequencies: Dict[int, int] = {}
                                for value in common_amounts:
                                    frequencies[value] = frequencies.get(value, 0) + 1
                                common_amount = min(
                                    frequencies,
                                    key=lambda value: (-frequencies[value], value),
                                )
                                if (
                                    common_amount > 0
                                    and candidate.metrics.total_diff_li % common_amount == 0
                                ):
                                    estimated_count = (
                                        candidate.metrics.total_diff_li // common_amount
                                    )
                                    estimate_hint = (
                                        f"；按常见单笔金额测算，差额相当于{estimated_count}笔，"
                                        "但不能确定具体是缺项、重复、补发还是退回"
                                    )
                        supporting_document = {
                            "工资": "工资表",
                            "奖金": "奖金发放表",
                            "员工报销": "员工报销清单",
                            "报销": "报销清单",
                            "供应商批付": "供应商付款清单",
                            "客户批收": "客户收款清单",
                            "社保公积金": "社保公积金明细",
                            "税费代扣": "税费申报及扣款明细",
                        }.get(
                            str(evidence.get("batch_category", "")),
                            "对应批量业务清单",
                        )
                        review_hint = (
                            f"{count_hint}；{amount_hint}{estimate_hint}；"
                            f"需结合{supporting_document}、银行回单和凭证逐项核查。"
                        )
                    else:
                        possible_side = "未定位"
                    candidate.evidence.update({
                        "batch_bank_count": len(candidate.bank_idxs),
                        "batch_journal_count": len(candidate.journal_idxs),
                        "batch_count_difference": count_difference,
                        "batch_difference_li": candidate.metrics.total_diff_li,
                        "batch_possible_missing_side": (
                            possible_side
                        ),
                        "batch_bank_source_rows": tuple(
                            int(self.bank.at[index, "original_file_row"])
                            if "original_file_row" in self.bank.columns
                            else int(index)
                            for index in candidate.bank_idxs
                        ),
                        "batch_journal_source_rows": tuple(
                            int(self.journal.at[index, "original_file_row"])
                            if "original_file_row" in self.journal.columns
                            else int(index)
                            for index in candidate.journal_idxs
                        ),
                        "batch_review_hint": review_hint,
                    })
                if candidate is not None and complete_voucher_group:
                    period = journal_group_key[1]
                    voucher_word = str(
                        self.journal.at[journal_idxs[0], "voucher_word"]
                    ).strip() if "voucher_word" in self.journal.columns else ""
                    voucher = str(
                        self.journal.at[journal_idxs[0], "voucher_no"]
                    ).strip()
                    cross_side_basis = candidate.evidence.get("business_basis", "")
                    candidate.evidence["business_basis"] = (
                        f"日记账同期间完整凭证 {period} {voucher_word}{voucher}；{cross_side_basis}"
                    ).rstrip("；")
                    self._score_existing_candidate(candidate)
            self._record_candidate_pool(
                "business_group",
                tuple(int(index) for index in bank_idxs),
                examined_count,
                evaluated_count,
                truncated=examined_count > evaluated_count,
            )
        self._add_atomic_journal_voucher_candidates(groups["journal"])
        self._commit_if_standalone()

    def _add_atomic_journal_voucher_candidates(
        self,
        journal_groups: List[Tuple[Any, Tuple[int, ...]]],
    ) -> None:
        """把完整日记账凭证作为一个候选单位，金额有差异时也不拆行凑数。"""
        performance_li = PrecisionEngine.to_integer_li(
            self.config.performance_materiality
        )
        self._reset_candidate_search_stage("atomic_voucher")
        bank_buckets: Dict[Tuple[int, pd.Timestamp], Dict[str, Any]] = {}
        bank_by_profile_key = self._global_profile_indexes["bank"]
        bank_dates_by_sign: Dict[int, Set[pd.Timestamp]] = {
            -1: set(),
            0: set(),
            1: set(),
        }
        for row in self.bank[["date", "amount_decimal"]].itertuples():
            amount = int(row.amount_decimal)
            sign = 1 if amount > 0 else -1 if amount < 0 else 0
            date = pd.Timestamp(row.date)
            bucket = bank_buckets.setdefault(
                (sign, date),
                {"entries": []},
            )
            bucket["entries"].append((amount, int(row.Index)))
            bank_dates_by_sign[sign].add(date)
        for bucket in bank_buckets.values():
            bucket["entries"].sort(
                key=lambda item: (
                    item[0],
                    self._stable_row_order_key("bank", item[1]),
                )
            )
            bucket["amounts"] = [item[0] for item in bucket["entries"]]
        ordered_bank_dates = {
            sign: sorted(values)
            for sign, values in bank_dates_by_sign.items()
        }
        for group_key, journal_idxs in journal_groups:
            if group_key[0] != "凭证" or len(journal_idxs) <= 1:
                continue
            journal_rows = [
                self._business_rows["journal"][index]
                for index in journal_idxs
            ]
            journal_dates = [row["date"] for row in journal_rows]
            journal_date_min = min(journal_dates)
            journal_date_max = max(journal_dates)
            if (
                journal_date_max - journal_date_min
            ).days > self.config.dfs_date_window:
                continue
            journal_signs = {row["sign"] for row in journal_rows}
            if len(journal_signs) != 1:
                continue
            journal_sign = next(iter(journal_signs))
            journal_total = sum(
                int(self.journal.at[index, "amount_decimal"])
                for index in journal_idxs
            )
            ranked: List[Tuple[int, int, int, str, int, Dict[str, Any]]] = []
            date_min = journal_date_max - timedelta(
                days=self.config.dfs_date_window
            )
            date_max = journal_date_min + timedelta(
                days=self.config.dfs_date_window
            )
            candidate_dates = ordered_bank_dates[journal_sign]
            start = bisect.bisect_left(candidate_dates, date_min)
            end = bisect.bisect_right(candidate_dates, date_max)
            amount_min = journal_total - performance_li
            amount_max = journal_total + performance_li
            eligible_bank_indexes: Set[int] = set()
            for date_position in range(start, end):
                bank_date = candidate_dates[date_position]
                bucket = bank_buckets[(journal_sign, bank_date)]
                amount_start = bisect.bisect_left(
                    bucket["amounts"],
                    amount_min,
                )
                amount_end = bisect.bisect_right(
                    bucket["amounts"],
                    amount_max,
                )
                for position in range(amount_start, amount_end):
                    eligible_bank_indexes.add(bucket["entries"][position][1])

            # 金额窗之外只补入共同编号，或“明确对方 + 同一具体摘要”命中的银行记录。
            # 同一凭证的重复键先去重，再用索引交集缩小范围，不能对所有命中行反复扫描整张凭证。
            journal_profile_keys = {
                profile_key
                for journal_row in journal_rows
                for profile_key in self._profile_pool_keys(journal_row)
            }
            identifier_keys = {
                key
                for key in journal_profile_keys
                if key and key[0] in {"业务编号", "交易流水"}
            }
            identifier_bank_indexes: Set[int] = set()
            for profile_key in identifier_keys:
                identifier_bank_indexes.update(
                    bank_by_profile_key.get(profile_key, set())
                )

            journal_summaries = {
                str(row.get("summary", "")).strip()
                for row in journal_rows
                if str(row.get("summary", "")).strip()
            }
            party_values: Dict[str, Set[str]] = {}
            for row in journal_rows:
                for category, values in row.get("parties", {}).items():
                    party_values.setdefault(str(category), set()).update(values)
            party_bank_indexes: Set[int] = set()
            for category, values in party_values.items():
                if len(values) != 1:
                    continue
                party_bank_indexes.update(
                    bank_by_profile_key.get(
                        ("对方", category, next(iter(values))),
                        set(),
                    )
                )
            if len(journal_summaries) == 1 and party_bank_indexes:
                summary_bank_indexes = set(
                    bank_by_profile_key.get(
                        ("摘要", next(iter(journal_summaries))),
                        set(),
                    )
                )
                party_summary_bank_indexes = (
                    party_bank_indexes & summary_bank_indexes
                )
            else:
                party_summary_bank_indexes = set()
            strong_bank_indexes = (
                identifier_bank_indexes | party_summary_bank_indexes
            )
            eligible_bank_indexes.update(strong_bank_indexes)

            cheap_ranked: List[Tuple[int, int, int, str, int]] = []
            for bank_index in eligible_bank_indexes:
                bank_amount = int(self.bank.at[bank_index, "amount_decimal"])
                bank_date = pd.Timestamp(self.bank.at[bank_index, "date"])
                bank_sign = 1 if bank_amount > 0 else -1 if bank_amount < 0 else 0
                if (
                    bank_sign != journal_sign
                    or bank_date < date_min
                    or bank_date > date_max
                ):
                    continue
                difference = abs(bank_amount - journal_total)
                date_span = int(
                    (
                        max(journal_date_max, bank_date)
                        - min(journal_date_min, bank_date)
                    ).days
                )
                cheap_ranked.append((
                    0 if bank_index in identifier_bank_indexes else (
                        1 if bank_index in party_summary_bank_indexes else 2
                    ),
                    difference,
                    date_span,
                    self._stable_row_order_key("bank", bank_index)[0],
                    bank_index,
                ))

            max_candidates = max(0, int(self.config.max_candidates))
            evidence_evaluation_limit = max_candidates * 4
            for _, difference, date_span, stable_hash, bank_index in sorted(
                cheap_ranked,
            )[:evidence_evaluation_limit]:
                bank_amount = int(self.bank.at[bank_index, "amount_decimal"])
                bank_row = self._business_rows["bank"][bank_index]
                evidence = business_evidence([bank_row], journal_rows)
                strong_scope_evidence = bool(
                    evidence.get("shared_business_id")
                    or (
                        evidence.get("same_party_group")
                        and evidence.get("specific_summary_group")
                    )
                )
                outside_amount_window = not (
                    amount_min <= bank_amount <= amount_max
                )
                if outside_amount_window:
                    if (
                        not strong_scope_evidence
                        or evidence["business_conflicts"]
                    ):
                        continue
                else:
                    blocking_conflicts = set(evidence["business_conflicts"]) - {
                        "业务编号或批次不同"
                    }
                    if blocking_conflicts:
                        continue
                ranked.append((
                    0 if strong_scope_evidence else 1,
                    difference,
                    date_span,
                    stable_hash,
                    bank_index,
                    evidence,
                ))

            retained = min(len(ranked), max_candidates)
            self._record_candidate_pool(
                "atomic_voucher",
                tuple(int(index) for index in journal_idxs),
                len(cheap_ranked),
                retained,
            )
            for _, difference, _, _, bank_index, group_evidence in sorted(
                ranked,
                key=lambda item: item[:5],
            )[:retained]:
                cross_side_evidence = self._has_strong_cross_side_evidence(
                    group_evidence
                ) or bool(group_evidence.get("same_party_group"))
                existing = next(
                    (
                        candidate for candidate in self.candidates
                        if candidate.match_type == "business_group"
                        and candidate.bank_idxs == (bank_index,)
                        and candidate.journal_idxs == tuple(journal_idxs)
                    ),
                    None,
                )
                candidate = existing or self._add_candidate(
                    [bank_index],
                    journal_idxs,
                    "business_group",
                    "完整凭证组",
                    resolves_full_group=cross_side_evidence,
                    is_rule_matched=True,
                    complete_business_group=cross_side_evidence,
                    represents_full_observed_group=True,
                    business_group_kind="凭证",
                    atomic_voucher_group=True,
                    batch_difference=bool(difference),
                    _precomputed_business_evidence=group_evidence,
                )
                if candidate is None:
                    continue
                period = group_key[1]
                voucher_word = group_key[2]
                voucher = group_key[3]
                prior_basis = candidate.evidence.get("business_basis", "")
                aggregate_party_hint = str(
                    candidate.evidence.get("aggregate_party_hierarchy_hint", "")
                ).strip()
                difference_hint = (
                    "完整凭证与银行记录存在金额差异；保留整张凭证核查，"
                    "不得抽取部分分录凑成相等关系。"
                    if difference
                    else ""
                )
                combined_review_hint = "；".join(
                    hint.rstrip("；。")
                    for hint in (difference_hint, aggregate_party_hint)
                    if hint
                )
                if combined_review_hint:
                    combined_review_hint += "。"
                candidate.evidence.update({
                    "business_basis": (
                        f"日记账同期间完整凭证 {period} {voucher_word}{voucher}；"
                        f"{prior_basis}"
                    ).rstrip("；"),
                    "resolves_full_group": cross_side_evidence,
                    "complete_business_group": cross_side_evidence,
                    "represents_full_observed_group": True,
                    "business_group_kind": "凭证",
                    "atomic_voucher_group": True,
                    "batch_difference": bool(difference),
                    "batch_difference_li": difference,
                    "batch_bank_count": 1,
                    "batch_journal_count": len(journal_idxs),
                    "batch_review_hint": combined_review_hint,
                })
                self._score_existing_candidate(candidate)

    def match_exact_1to1(self) -> None:
        """为同日同金额记录生成一对一候选，不在本阶段抢占记录。"""
        self._reset_candidate_search_stage("exact")
        b_groups = self.bank.groupby(["date", "amount_decimal"]).groups
        j_groups = self.journal.groupby(["date", "amount_decimal"]).groups
        common_keys = set(b_groups.keys()) & set(j_groups.keys())
        for key in sorted(common_keys, key=lambda item: (item[0], item[1])):
            journal_indexes = sorted(
                (int(index) for index in j_groups[key]),
                key=lambda index: self._stable_row_order_key("journal", index),
            )
            profile_index = self._build_profile_pool_index(
                journal_indexes,
                "journal",
            )
            for bank_index in sorted(
                (int(index) for index in b_groups[key]),
                key=lambda index: self._stable_row_order_key("bank", index),
            ):
                pool = self._stable_profile_pool(
                    self._business_rows["bank"][bank_index],
                    journal_indexes,
                    "journal",
                    self.config.max_candidates,
                    profile_index,
                )
                self._record_candidate_pool(
                    "exact",
                    bank_index,
                    len(journal_indexes),
                    len(pool),
                )
                ranked = []
                for journal_index in pool:
                    evidence = score_text_fields(
                        self._row_text_fields(self.bank, bank_index),
                        self._row_text_fields(self.journal, journal_index),
                    )
                    ranked.append((-evidence.local_score, journal_index))
                for _, journal_index in sorted(ranked):
                    self._add_candidate(
                        [bank_index],
                        [journal_index],
                        "exact_1to1",
                        "精确",
                    )

    def match_tolerance(self) -> None:
        """生成日期容差和明显微小金额差异候选。"""
        self._reset_candidate_search_stage("tolerance")
        trivial_li = PrecisionEngine.to_integer_li(
            self.config.clearly_trivial_threshold
        )
        buckets: Dict[Tuple[int, pd.Timestamp], Dict[str, Any]] = {}
        days_by_sign: Dict[int, Set[pd.Timestamp]] = {-1: set(), 0: set(), 1: set()}
        for row in self.journal[["amount_decimal", "date"]].itertuples():
            amount = int(row.amount_decimal)
            sign = 1 if amount > 0 else -1 if amount < 0 else 0
            date = pd.Timestamp(row.date)
            index = int(row.Index)
            bucket = buckets.setdefault(
                (sign, date),
                {"entries": [], "by_key": {}},
            )
            bucket["entries"].append((amount, index))
            for profile_key in self._profile_pool_keys(
                self._business_rows["journal"][index]
            ):
                bucket["by_key"].setdefault(profile_key, []).append(
                    (amount, index)
                )
            days_by_sign[sign].add(date)
        for bucket in buckets.values():
            bucket["entries"].sort(
                key=lambda item: (
                    item[0],
                    self._stable_row_order_key("journal", item[1]),
                )
            )
            bucket["amounts"] = [item[0] for item in bucket["entries"]]
            for profile_key, entries in bucket["by_key"].items():
                entries.sort(
                    key=lambda item: (
                        item[0],
                        self._stable_row_order_key("journal", item[1]),
                    )
                )
                bucket["by_key"][profile_key] = (
                    entries,
                    [item[0] for item in entries],
                )
        ordered_days = {
            sign: sorted(values)
            for sign, values in days_by_sign.items()
        }

        def nearby_entries(
            entries: List[Tuple[int, int]],
            amounts: List[int],
            lower: int,
            upper: int,
            target: int,
            limit: int,
            skip_exact: bool,
            excluded: Set[int],
        ) -> List[Tuple[int, int]]:
            """从金额区间两侧向外取最近记录，读取量受 limit 约束。"""
            left_bound = bisect.bisect_left(amounts, lower)
            right_bound = bisect.bisect_right(amounts, upper)
            if skip_exact:
                left = bisect.bisect_left(
                    amounts,
                    target,
                    left_bound,
                    right_bound,
                ) - 1
                right = bisect.bisect_right(
                    amounts,
                    target,
                    left_bound,
                    right_bound,
                )
            else:
                split = bisect.bisect_left(
                    amounts,
                    target,
                    left_bound,
                    right_bound,
                )
                left = split - 1
                right = split
            chosen: List[Tuple[int, int]] = []
            while len(chosen) < limit and (left >= left_bound or right < right_bound):
                left_diff = (
                    abs(entries[left][0] - target)
                    if left >= left_bound else None
                )
                right_diff = (
                    abs(entries[right][0] - target)
                    if right < right_bound else None
                )
                if right_diff is not None and (
                    left_diff is None or right_diff <= left_diff
                ):
                    item = entries[right]
                    right += 1
                else:
                    item = entries[left]
                    left -= 1
                if skip_exact and item[0] == target:
                    continue
                if item[1] in excluded:
                    continue
                chosen.append(item)
            return chosen

        for bank_row in self.bank.sort_index().itertuples():
            bank_index = int(bank_row.Index)
            bank_amount = int(bank_row.amount_decimal)
            bank_date = pd.Timestamp(bank_row.date)
            sign = 1 if bank_amount > 0 else -1 if bank_amount < 0 else 0
            days = ordered_days[sign]
            minimum_date = bank_date - timedelta(days=self.config.tolerance_days)
            maximum_date = bank_date + timedelta(days=self.config.tolerance_days)
            day_start = bisect.bisect_left(days, minimum_date)
            day_end = bisect.bisect_right(days, maximum_date)
            candidate_days = sorted(
                days[day_start:day_end],
                key=lambda value: (abs((bank_date - value).days), value),
            )
            lower = bank_amount - trivial_li
            upper = bank_amount + trivial_li
            examined = 0
            for day in candidate_days:
                bucket = buckets[(sign, day)]
                start = bisect.bisect_left(bucket["amounts"], lower)
                end = bisect.bisect_right(bucket["amounts"], upper)
                count = end - start
                if day == bank_date:
                    count -= (
                        bisect.bisect_right(bucket["amounts"], bank_amount)
                        - bisect.bisect_left(bucket["amounts"], bank_amount)
                    )
                examined += max(0, count)
            if examined <= 0:
                continue

            limit = max(0, int(self.config.max_candidates))
            selected: List[Tuple[int, int, int]] = []
            seen: Set[int] = set()
            source_keys = self._profile_pool_keys(
                self._business_rows["bank"][bank_index]
            )

            def add_from_bucket(
                day: pd.Timestamp,
                entries: List[Tuple[int, int]],
                amounts: List[int],
            ) -> None:
                remaining = limit - len(selected)
                if remaining <= 0:
                    return
                for journal_amount, journal_index in nearby_entries(
                    entries,
                    amounts,
                    lower,
                    upper,
                    bank_amount,
                    remaining,
                    day == bank_date,
                    seen,
                ):
                    seen.add(journal_index)
                    selected.append(
                        (
                            abs(bank_amount - journal_amount),
                            abs((bank_date - day).days),
                            journal_index,
                        )
                    )

            for profile_key in source_keys:
                for day in candidate_days:
                    bucket = buckets[(sign, day)]
                    entries_and_amounts = bucket["by_key"].get(profile_key)
                    if entries_and_amounts is not None:
                        add_from_bucket(
                            day,
                            entries_and_amounts[0],
                            entries_and_amounts[1],
                        )
                    if len(selected) >= limit:
                        break
                if len(selected) >= limit:
                    break
            if len(selected) < limit:
                for day in candidate_days:
                    bucket = buckets[(sign, day)]
                    add_from_bucket(
                        day,
                        bucket["entries"],
                        bucket["amounts"],
                    )
                    if len(selected) >= limit:
                        break

            self._record_candidate_pool(
                "tolerance",
                bank_index,
                examined,
                len(selected),
            )
            ranked = []
            for amount_difference, date_difference, journal_index in selected:
                evidence = score_text_fields(
                    self._row_text_fields(self.bank, bank_index),
                    self._row_text_fields(self.journal, journal_index),
                )
                ranked.append(
                    (
                        -evidence.local_score,
                        amount_difference,
                        date_difference,
                        journal_index,
                    )
                )
            for _, amount_difference, _, journal_index in sorted(ranked):
                self._add_candidate(
                    [int(bank_index)],
                    [journal_index],
                    (
                        "tolerance_date"
                        if amount_difference == 0
                        else "amount_difference"
                    ),
                    (
                        "容差"
                        if amount_difference == 0
                        else "金额差异"
                    ),
                )

    def match_batch_aggregation(self) -> None:
        """同一天、同金额达到门槛且文字明确时生成批量候选。"""
        if self.bank.empty or self.journal.empty:
            return
        batch_keywords = ('工资', '代发', '批量', '发放', '薪资', '薪酬')
        grouped = self.bank.groupby(
            [self.bank["date"].dt.normalize(), "amount_decimal"],
            sort=True,
        )
        for (group_date, _), group in grouped:
            if len(group) < self.config.batch_min_count:
                continue
            bank_texts = [
                self._rule_search_text(self.bank, int(index)).strip()
                for index in group.index
            ]
            non_empty_texts = [text for text in bank_texts if text]
            if not non_empty_texts:
                continue
            if not all(
                any(keyword in text for keyword in batch_keywords)
                for text in non_empty_texts
            ):
                continue

            total_amount = group["amount_decimal"].sum()
            journal_candidates = self.journal[
                self.journal["amount_decimal"] == total_amount
            ]
            for journal_index, journal_row in journal_candidates.iterrows():
                date_difference = abs(
                    (pd.Timestamp(journal_row["date"]).normalize() - group_date).days
                )
                if date_difference > 3:
                    continue
                journal_text = self._rule_search_text(
                    self.journal,
                    int(journal_index),
                )
                if not any(
                    keyword in journal_text for keyword in batch_keywords
                ):
                    continue
                self._add_candidate(
                    [int(index) for index in group.index],
                    [int(journal_index)],
                    "batch_aggregation",
                    "聚合",
                    is_rule_matched=True,
                    batch_count=len(group),
                    resolves_full_group=True,
                    complete_business_group=True,
                    represents_full_observed_group=True,
                )
        self._commit_if_standalone()

    def match_continuous_summary_groups(self) -> None:
        """为物理相邻、标准化摘要一致的完整连续组生成双向候选。"""
        self._match_continuous_summary_side("bank", self.bank, self.journal)
        self._match_continuous_summary_side("journal", self.journal, self.bank)
        self._commit_if_standalone()

    def _match_continuous_summary_side(
        self,
        source_side: str,
        source: pd.DataFrame,
        target: pd.DataFrame,
    ) -> None:
        if (
            source.empty
            or target.empty
            or "original_idx" not in source.columns
        ):
            return

        target_by_summary: Dict[str, List[int]] = {}
        for target_index, target_row in target.iterrows():
            summary = normalize_summary(target_row.get("summary", ""))
            if summary:
                target_by_summary.setdefault(summary, []).append(int(target_index))

        def add_run(run: List[int], summary: str) -> None:
            if len(run) < 2:
                return
            source_amounts = source.loc[run, "amount_decimal"].tolist()
            source_dates = [pd.Timestamp(value) for value in source.loc[run, "date"]]
            source_rows = tuple(int(value) for value in source.loc[run, "original_idx"])
            for target_index in target_by_summary.get(summary, []):
                target_amount = int(target.at[target_index, "amount_decimal"])
                all_amounts = source_amounts + [target_amount]
                if not self.config.allow_zero_match and any(
                    amount == 0 for amount in all_amounts
                ):
                    continue
                target_date = pd.Timestamp(target.at[target_index, "date"])
                date_span = max(source_dates + [target_date]) - min(
                    source_dates + [target_date]
                )
                if date_span.days > self.config.dfs_date_window:
                    continue
                if source_side == "bank":
                    bank_amounts, journal_amounts = source_amounts, [target_amount]
                    bank_idxs, journal_idxs = run, [target_index]
                else:
                    bank_amounts, journal_amounts = [target_amount], source_amounts
                    bank_idxs, journal_idxs = [target_index], run
                if not self._total_structure_matches(bank_amounts, journal_amounts):
                    continue
                self._add_candidate(
                    bank_idxs,
                    journal_idxs,
                    "continuous_summary_group",
                    "连续摘要整组",
                    is_rule_matched=True,
                    source_side=source_side,
                    normalized_summary=summary,
                    group_count=len(run),
                    source_rows=source_rows,
                )

        run: List[int] = []
        run_summary = ""
        previous_row: Optional[int] = None
        for source_index, row in source.sort_values("original_idx").iterrows():
            summary = normalize_summary(row.get("summary", ""))
            source_row = pd.to_numeric(row.get("original_idx"), errors="coerce")
            if pd.isna(source_row):
                add_run(run, run_summary)
                run, run_summary, previous_row = [], "", None
                continue
            original_row = int(source_row)
            if (
                run
                and summary
                and summary == run_summary
                and previous_row is not None
                and original_row == previous_row + 1
            ):
                run.append(int(source_index))
            else:
                add_run(run, run_summary)
                run = [int(source_index)] if summary else []
                run_summary = summary
            previous_row = original_row
        add_run(run, run_summary)

    def match_dfs_combinations(self) -> None:
        self._combination_covered = {"bank": set(), "journal": set()}
        for side, groups in self._atomic_groups.items():
            for atomic_group in groups:
                self._combination_covered[side].update(atomic_group)
        for candidate in self.candidates:
            if (candidate.evidence.get("resolves_full_group")
                    and candidate.evidence.get("business_strength", 0) >= 2
                    and (candidate.metrics.total_diff_li == 0
                         or candidate.evidence.get("complete_business_id")
                         or candidate.evidence.get("complete_business_group"))
                    and not has_business_conflict(candidate)):
                self._combination_covered["bank"].update(candidate.bank_idxs)
                self._combination_covered["journal"].update(candidate.journal_idxs)
        self._reset_candidate_search_stage("generic_combination")
        self._combination_source_sets = {
            "generic": set(),
            "fully_searched": set(),
            "truncated": set(),
            "depth_limited": set(),
            "node_budget_exhausted": set(),
            "task_timeout": set(),
            "global_timeout_unprocessed": set(),
            "worker_failure": set(),
            "budget_exhausted": set(),
        }
        global_limit = max(
            0.0,
            float(self.config.combination_global_time_limit_seconds),
        )
        self._combination_global_deadline = time.monotonic() + global_limit
        self.run_parameters["combination_search"] = {
            "business_group_bank_rows": len(self._combination_covered["bank"]),
            "business_group_journal_rows": len(self._combination_covered["journal"]),
            "generic_source_rows": 0,
            "fully_searched_source_rows": 0,
            "truncated_source_rows": 0,
            "depth_limited_source_rows": 0,
            "node_budget_exhausted_source_rows": 0,
            "task_timeout_source_rows": 0,
            "global_timeout_unprocessed_source_rows": 0,
            "worker_failure_source_rows": 0,
            "candidate_limit": min(self.config.max_candidates, 30),
            "search_budget": {
                "generic_candidate_limit": min(self.config.max_candidates, 30),
                "preprocessing_evidence_limit_per_source": (
                    min(self.config.max_candidates, 30) * 4
                ),
                "max_depth": self.config.max_dfs_depth,
                "exact_split_max_depth": 8,
                "node_limit_per_source": self.config.combination_node_limit_per_source,
                "task_timeout_seconds": self.config.combination_task_timeout_seconds,
                "global_time_limit_seconds": self.config.combination_global_time_limit_seconds,
            },
            "budget_exhausted_source_rows": 0,
        }
        self._dfs_one_to_many('bank', 'journal')
        self._dfs_one_to_many('journal', 'bank')
        self._sync_combination_search_counts()
        self._commit_if_standalone()

    def _sync_combination_search_counts(self) -> None:
        """把内部去重集合汇总为可序列化的运行统计。"""
        search = self.run_parameters.get("combination_search")
        source_sets = getattr(self, "_combination_source_sets", None)
        if search is None or source_sets is None:
            return
        for set_name, field_name in (
            ("generic", "generic_source_rows"),
            ("fully_searched", "fully_searched_source_rows"),
            ("truncated", "truncated_source_rows"),
            ("depth_limited", "depth_limited_source_rows"),
            ("node_budget_exhausted", "node_budget_exhausted_source_rows"),
            ("task_timeout", "task_timeout_source_rows"),
            ("global_timeout_unprocessed", "global_timeout_unprocessed_source_rows"),
            ("worker_failure", "worker_failure_source_rows"),
            ("budget_exhausted", "budget_exhausted_source_rows"),
        ):
            search[field_name] = len(source_sets[set_name])

    def _record_combination_outcome(
        self,
        source_type: str,
        outcome: _CombinationTaskResult,
    ) -> None:
        """合并一个来源的搜索结果，并按来源去重记录各类限制。"""
        source_key = (source_type, int(outcome.source_idx))
        source_sets = self._combination_source_sets
        if outcome.candidate_truncated:
            source_sets["truncated"].add(source_key)
            source_sets["budget_exhausted"].add(source_key)
        if outcome.depth_limited:
            source_sets["depth_limited"].add(source_key)
            source_sets["budget_exhausted"].add(source_key)
        if outcome.exhaustion_reason == "node_limit":
            source_sets["node_budget_exhausted"].add(source_key)
            source_sets["budget_exhausted"].add(source_key)
        elif outcome.exhaustion_reason == "task_timeout":
            source_sets["task_timeout"].add(source_key)
            source_sets["budget_exhausted"].add(source_key)
        if outcome.fully_searched:
            source_sets["fully_searched"].add(source_key)
        self._sync_combination_search_counts()

    @staticmethod
    def _stop_process_pool(executor: ProcessPoolExecutor) -> None:
        """立即终止未完成的组合任务；Python 3.14 优先使用原生终止接口。"""
        terminate = getattr(executor, "terminate_workers", None)
        try:
            if callable(terminate):
                terminate()
            else:
                executor.shutdown(wait=False, cancel_futures=True)
        except (BrokenProcessPool, RuntimeError):
            try:
                executor.shutdown(wait=False, cancel_futures=True)
            except (BrokenProcessPool, RuntimeError):
                pass

    def _dfs_one_to_many(self, source_type: str, target_type: str) -> None:
        if source_type == 'bank':
            sources = self.bank
            targets_df = self.journal
        else:
            sources = self.journal
            targets_df = self.bank

        covered = getattr(self, "_combination_covered", {"bank": set(), "journal": set()})
        sources = sources.loc[~sources.index.isin(covered[source_type])]
        targets_df = targets_df.loc[~targets_df.index.isin(covered[target_type])]
            
        if sources.empty or targets_df.empty: return

        tgt_view = targets_df.reset_index()
        tgt_view["_stable_content_order"] = [
            self._row_content_hash(targets_df, int(index))
            for index in tgt_view["index"]
        ]
        tgt_view = tgt_view.sort_values(
            ["date", "_stable_content_order", "index"],
            kind="mergesort",
        )
        tgt_dates = tgt_view['date'].values.astype('datetime64[ns]')
        tgt_indexes = tgt_view['index'].values
        tgt_amounts = tgt_view['amount_decimal'].values
        tgt_view_dict = []
        for i in range(len(tgt_view)):
            idx = tgt_indexes[i]
            tgt_view_dict.append({
                'index': idx,
                'date': tgt_view.iloc[i]['date'],
                'amount_decimal': tgt_amounts[i],
                'matched': False,
                'stable_key': tgt_view.iloc[i]['_stable_content_order'],
            })

        def new_bucket() -> Dict[str, List[Any]]:
            return {"items": [], "dates": []}

        target_sign_buckets: Dict[int, Dict[str, List[Any]]] = {
            -1: new_bucket(),
            0: new_bucket(),
            1: new_bucket(),
        }
        target_amount_buckets: Dict[int, Dict[str, List[Any]]] = {}
        target_profile_buckets: Dict[
            Tuple[Tuple[Any, ...], int], Dict[str, List[Any]]
        ] = {}
        target_amount_values_by_sign: Dict[int, Set[int]] = {
            -1: set(),
            0: set(),
            1: set(),
        }
        for item in tgt_view_dict:
            amount = int(item["amount_decimal"])
            sign = 1 if amount > 0 else -1 if amount < 0 else 0
            date = pd.Timestamp(item["date"])
            index = int(item["index"])
            for bucket in (
                target_sign_buckets[sign],
                target_amount_buckets.setdefault(amount, new_bucket()),
            ):
                bucket["items"].append(item)
                bucket["dates"].append(date)
            target_amount_values_by_sign[sign].add(amount)
            for profile_key in self._profile_pool_keys(
                self._business_rows[target_type][index]
            ):
                bucket = target_profile_buckets.setdefault(
                    (profile_key, sign), new_bucket()
                )
                bucket["items"].append(item)
                bucket["dates"].append(date)
        ordered_amount_values_by_sign = {
            sign: sorted(values)
            for sign, values in target_amount_values_by_sign.items()
        }

        def bucket_bounds(
            bucket: Dict[str, List[Any]],
            date_min: pd.Timestamp,
            date_max: pd.Timestamp,
        ) -> Tuple[int, int]:
            dates = bucket["dates"]
            return (
                bisect.bisect_left(dates, date_min),
                bisect.bisect_right(dates, date_max),
            )

        def bounded_bucket_items(
            bucket: Optional[Dict[str, List[Any]]],
            date_min: pd.Timestamp,
            date_max: pd.Timestamp,
            center: pd.Timestamp,
            limit: int,
        ) -> List[Dict[str, Any]]:
            if not bucket or limit <= 0:
                return []
            start, end = bucket_bounds(bucket, date_min, date_max)
            if end - start <= limit:
                return list(bucket["items"][start:end])
            middle = bisect.bisect_left(
                bucket["dates"], center, start, end
            )
            left = max(start, middle - limit // 2)
            right = min(end, left + limit)
            left = max(start, right - limit)
            return list(bucket["items"][left:right])

        tasks: List[Tuple[Tuple[Any, ...], Tuple[str, int]]] = []
        source_indices = sorted(
            (int(index) for index in sources.index),
            key=lambda index: self._stable_row_order_key(source_type, index),
        )
        date_window = self.config.dfs_date_window

        def mark_preprocessing_timeout(
            remaining_indices: List[int],
        ) -> None:
            """候选窗口预处理同样受全局预算约束，并披露尚未处理的来源行。"""
            pending_keys = [source_key for _, source_key in tasks]
            pending_keys.extend(
                (source_type, int(source_index))
                for source_index in remaining_indices
            )
            for source_key in pending_keys:
                self._combination_source_sets[
                    "global_timeout_unprocessed"
                ].add(source_key)
                self._combination_source_sets["budget_exhausted"].add(source_key)
            self._sync_combination_search_counts()
        
        source_dates = sources['date'].to_dict()
        source_amounts = sources['amount_decimal'].to_dict()
        
        for source_position, s_idx in enumerate(source_indices):
            if self.stopping:
                return
            if time.monotonic() >= self._combination_global_deadline:
                mark_preprocessing_timeout(source_indices[source_position:])
                return
            target_val = source_amounts[s_idx]
            target_val_int = int(target_val)
            s_date = pd.Timestamp(source_dates[s_idx])

            date_min = s_date - timedelta(days=date_window)
            date_max = s_date + timedelta(days=date_window)
            if self.config.allow_mixed_sign or target_val_int == 0:
                allowed_signs = (-1, 0, 1)
            else:
                allowed_signs = (1 if target_val_int > 0 else -1,)
            raw_window_count = sum(
                end - start
                for sign in allowed_signs
                for start, end in [bucket_bounds(
                    target_sign_buckets[sign], date_min, date_max
                )]
            )
            exact_count = 0
            if target_val_int != 0:
                exact_bucket = target_amount_buckets.get(target_val_int)
                if exact_bucket:
                    exact_start, exact_end = bucket_bounds(
                        exact_bucket, date_min, date_max
                    )
                    exact_count = exact_end - exact_start
            exact_single_available = bool(target_val_int != 0 and exact_count)
            candidate_pool_size = max(
                0,
                raw_window_count - (exact_count if exact_single_available else 0),
            )
            source_key = (source_type, int(s_idx))
            candidate_limit = max(
                0,
                int(self.run_parameters["combination_search"]["candidate_limit"]),
            )
            if candidate_pool_size < 2:
                if raw_window_count >= 2:
                    self._combination_source_sets["generic"].add(source_key)
                    self._combination_source_sets["fully_searched"].add(
                        source_key
                    )
                    self._record_candidate_pool(
                        "generic_combination",
                        source_key,
                        candidate_pool_size,
                        candidate_pool_size,
                        truncated=False,
                    )
                    self._sync_combination_search_counts()
                continue

            evidence_limit = candidate_limit * 4
            source_profile = self._business_rows[source_type][int(s_idx)]
            pooled: Dict[int, Tuple[int, Dict[str, Any]]] = {}

            def add_pool_items(
                items: List[Dict[str, Any]],
                strength: int,
            ) -> None:
                for item in items:
                    amount = int(item["amount_decimal"])
                    if exact_single_available and amount == target_val_int:
                        continue
                    index = int(item["index"])
                    prior = pooled.get(index)
                    if prior is None or strength > prior[0]:
                        pooled[index] = (strength, item)

            profile_strength = {
                "业务编号": 50,
                "交易流水": 50,
                "对方": 40,
                "批量用途": 40,
                "摘要": 30,
            }
            source_profile_keys = sorted(
                self._profile_pool_keys(source_profile),
                key=lambda key: (
                    -profile_strength.get(str(key[0]), 0),
                    tuple(map(str, key)),
                ),
            )
            for profile_key in source_profile_keys:
                strength = profile_strength.get(str(profile_key[0]), 0)
                for sign in allowed_signs:
                    add_pool_items(
                        bounded_bucket_items(
                            target_profile_buckets.get((profile_key, sign)),
                            date_min,
                            date_max,
                            s_date,
                            evidence_limit,
                        ),
                        strength,
                    )

            # 金额枢轴补足没有共同文字键的拆分候选；只查看固定数量的邻近金额桶。
            if target_val_int != 0 and evidence_limit:
                for part_count in range(
                    2,
                    min(max(2, int(self.config.max_dfs_depth)), 8) + 1,
                ):
                    pivot = target_val_int / part_count
                    for sign in allowed_signs:
                        amount_values = ordered_amount_values_by_sign[sign]
                        position = bisect.bisect_left(amount_values, pivot)
                        nearby_positions = range(
                            max(0, position - 2),
                            min(len(amount_values), position + 3),
                        )
                        for amount_position in nearby_positions:
                            amount = amount_values[amount_position]
                            if exact_single_available and amount == target_val_int:
                                continue
                            add_pool_items(
                                bounded_bucket_items(
                                    target_amount_buckets.get(amount),
                                    date_min,
                                    date_max,
                                    s_date,
                                    evidence_limit,
                                ),
                                10,
                            )
            for sign in allowed_signs:
                add_pool_items(
                    bounded_bucket_items(
                        target_sign_buckets[sign],
                        date_min,
                        date_max,
                        s_date,
                        evidence_limit,
                    ),
                    0,
                )

            ranked_pool = sorted(
                pooled.values(),
                key=lambda entry: (
                    -entry[0],
                    0
                    if (
                        self.config.allow_mixed_sign
                        or target_val_int == 0
                        or abs(int(entry[1]["amount_decimal"])) < abs(target_val_int)
                    )
                    else 1,
                    -abs(int(entry[1]["amount_decimal"])),
                    abs((pd.Timestamp(entry[1]["date"]) - s_date).days),
                    str(entry[1].get("stable_key", "")),
                    int(entry[1]["index"]),
                ),
            )[:evidence_limit]
            valid_pool: List[Dict[str, Any]] = []
            for _, item in ranked_pool:
                if time.monotonic() >= self._combination_global_deadline:
                    mark_preprocessing_timeout(source_indices[source_position:])
                    return
                target_profile = self._business_rows[target_type][int(item["index"])]
                evidence = business_evidence(
                    [source_profile] if source_type == "bank" else [target_profile],
                    [target_profile] if source_type == "bank" else [source_profile],
                )
                if not evidence["business_conflicts"]:
                    valid_pool.append(item)
            window_view_dict = valid_pool[:candidate_limit]
            retained_count = len(window_view_dict)
            search_truncated = bool(
                candidate_pool_size > len(ranked_pool)
                or len(valid_pool) > retained_count
            )
            source_sets = self._combination_source_sets
            source_sets["generic"].add(source_key)
            if search_truncated:
                source_sets["truncated"].add(source_key)
                source_sets["budget_exhausted"].add(source_key)
            self._record_candidate_pool(
                "generic_combination",
                source_key,
                candidate_pool_size,
                retained_count,
                truncated=search_truncated,
            )
            self._sync_combination_search_counts()
            if retained_count < 2:
                continue
            window_dates = np.array(
                [item["date"] for item in window_view_dict],
                dtype="datetime64[ns]",
            )
            filtered_targets_data = {'dates': window_dates, 'view_dict': window_view_dict}
            tasks.append((
                (
                    s_idx,
                    s_date,
                    target_val,
                    filtered_targets_data,
                    self._execution_config,
                ),
                source_key,
            ))

        if not tasks: return

        num_workers = min(os.cpu_count() or 4, len(tasks))
        parallel_threshold = max(
            self.PARALLEL_MIN_TASKS,
            num_workers * self.PARALLEL_TASKS_PER_WORKER,
        )
        use_parallel = (
            len(tasks) >= parallel_threshold
            and self._should_use_parallel(len(tasks))
        )
        results: List[Tuple[int, List[List[int]], str]] = []
        completed_sources: Set[Tuple[str, int]] = set()
        task_by_source = {source_key: task for task, source_key in tasks}
        progress_start = 54.0 if source_type == "bank" else 58.5
        progress_span = 4.5
        progress_seen: Set[Tuple[str, int]] = set()
        last_progress_log_at = time.monotonic()
        last_progress_log_count = 0
        last_progress_callback_value = progress_start

        def report_task_progress(source_key: Tuple[str, int]) -> None:
            nonlocal last_progress_log_at
            nonlocal last_progress_log_count
            nonlocal last_progress_callback_value
            progress_seen.add(source_key)
            processed = len(progress_seen)
            total = len(tasks)
            exact_progress = progress_start + processed / total * progress_span
            callback_progress = (
                exact_progress if processed == total else float(int(exact_progress))
            )
            if callback_progress > last_progress_callback_value:
                self._update_progress(callback_progress)
                last_progress_callback_value = callback_progress
            now = time.monotonic()
            log_interval = max(1, total // 20)
            if (
                processed == total
                or processed - last_progress_log_count >= log_interval
                or now - last_progress_log_at >= 5.0
            ):
                self._log(
                    f"智能组合匹配({source_type})已处理 "
                    f"{processed:,}/{total:,} 个来源任务"
                )
                last_progress_log_count = processed
                last_progress_log_at = now

        def accept_outcome(
            source_key: Tuple[str, int],
            outcome: _CombinationTaskResult,
        ) -> None:
            if source_key in completed_sources:
                return
            completed_sources.add(source_key)
            self._record_combination_outcome(source_type, outcome)
            if outcome.legacy_result:
                results.append(outcome.legacy_result)

        def mark_global_timeout(pending: List[Tuple[str, int]]) -> None:
            for source_key in pending:
                if source_key in completed_sources:
                    continue
                self._combination_source_sets[
                    "global_timeout_unprocessed"
                ].add(source_key)
                self._combination_source_sets["budget_exhausted"].add(source_key)
            self._sync_combination_search_counts()

        def run_serial_tasks(
            pending_keys: List[Tuple[str, int]],
        ) -> None:
            for position, source_key in enumerate(pending_keys):
                if source_key in completed_sources:
                    continue
                if self.stopping:
                    break
                remaining_global = self._combination_global_deadline - time.monotonic()
                if remaining_global <= 0:
                    mark_global_timeout(pending_keys[position:])
                    break
                task_timeout = min(
                    max(0.0, float(self.config.combination_task_timeout_seconds)),
                    remaining_global,
                )
                task = (*task_by_source[source_key], task_timeout)
                try:
                    outcome = _process_single_source_with_diagnostics(task)
                    accept_outcome(source_key, outcome)
                except Exception as exc:
                    self.exception_logger.record_exception(
                        f"dfs_{source_type}",
                        int(source_key[1]),
                        exc,
                    )
                    self._combination_source_sets["worker_failure"].add(source_key)
                    self._combination_source_sets["budget_exhausted"].add(source_key)
                    self._sync_combination_search_counts()
                finally:
                    report_task_progress(source_key)

        if use_parallel:
            self._log(f"智能组合匹配({source_type}): {len(tasks)} 个任务，使用 {num_workers} 个进程...")
            executor: Optional[ProcessPoolExecutor] = None
            futures: Dict[Any, Tuple[str, int]] = {}
            try:
                remaining_global = self._combination_global_deadline - time.monotonic()
                if remaining_global <= 0:
                    mark_global_timeout([source_key for _, source_key in tasks])
                    return
                executor = ProcessPoolExecutor(max_workers=num_workers)
                submitted_timeout = min(
                    max(0.0, float(self.config.combination_task_timeout_seconds)),
                    remaining_global,
                )
                futures = {
                    executor.submit(
                        _process_single_source_with_diagnostics,
                        (*task, submitted_timeout),
                    ): source_key
                    for task, source_key in tasks
                }
                for future in as_completed(futures, timeout=remaining_global):
                    if self.stopping:
                        break
                    source_key = futures[future]
                    try:
                        outcome = future.result()
                        accept_outcome(source_key, outcome)
                    except BrokenProcessPool:
                        raise
                    except Exception as exc:
                        self.exception_logger.record_exception(
                            f"dfs_{source_type}",
                            int(source_key[1]),
                            exc,
                        )
                    finally:
                        report_task_progress(source_key)
                if self.stopping:
                    self._stop_process_pool(executor)
                    executor = None
                else:
                    executor.shutdown(wait=True)
                    executor = None
                    pending = [
                        source_key
                        for _, source_key in tasks
                        if source_key not in completed_sources
                    ]
                    if pending:
                        run_serial_tasks(pending)
            except FuturesTimeoutError:
                for future, source_key in futures.items():
                    if source_key in completed_sources or not future.done():
                        continue
                    try:
                        accept_outcome(source_key, future.result())
                    except Exception:
                        pass
                pending = [
                    source_key
                    for _, source_key in tasks
                    if source_key not in completed_sources
                ]
                mark_global_timeout(pending)
                if executor is not None:
                    self._stop_process_pool(executor)
                    executor = None
            except Exception as exc:
                self._log(
                    f"智能组合匹配({source_type})后台进程不可用，"
                    "已在剩余预算内退回单线程继续"
                )
                self.exception_logger.record_exception(
                    f"dfs_{source_type}_parallel",
                    -1,
                    exc,
                )
                if executor is not None:
                    for future, source_key in futures.items():
                        if source_key in completed_sources or not future.done():
                            continue
                        try:
                            accept_outcome(source_key, future.result())
                        except Exception:
                            pass
                    self._stop_process_pool(executor)
                    executor = None
                if not self.stopping:
                    pending = [
                        source_key
                        for _, source_key in tasks
                        if source_key not in completed_sources
                    ]
                    run_serial_tasks(pending)
            finally:
                if executor is not None:
                    self._stop_process_pool(executor)
        else:
            self._log(f"智能组合匹配({source_type}): {len(tasks)} 个任务，单线程处理...")
            run_serial_tasks([source_key for _, source_key in tasks])

        source_order = {source_idx: order for order, source_idx in enumerate(source_indices)}

        results.sort(
            key=lambda item: (
                self._confidence_sort_key(item[2]),
                len(item[1][0]),
                source_order.get(item[0], len(source_order)),
                min(item[1][0]) if item[1] and item[1][0] else -1,
            )
        )
        
        for source_idx, matched_solutions, _confidence in results:
            for matched_idxs in matched_solutions:
                if source_type == 'bank':
                    bank_idxs = [int(source_idx)]
                    journal_idxs = [int(index) for index in matched_idxs]
                else:
                    bank_idxs = [int(index) for index in matched_idxs]
                    journal_idxs = [int(source_idx)]
                self._add_candidate(
                    bank_idxs,
                    journal_idxs,
                    "combination_dfs",
                    "组合",
                    combo_count=max(len(bank_idxs), len(journal_idxs)),
                )

    def match_monthly_total(self) -> None:
        self._match_total('month', self.MONTHLY_MAX_COUNT_DIFF_BASE, self.MONTHLY_MAX_COUNT_DIFF_RATIO, 'M')

    def match_daily_total(self) -> None:
        self._match_total('date', self.DAILY_MAX_COUNT_DIFF_BASE, self.DAILY_MAX_COUNT_DIFF_RATIO, None)

    def _match_total(self, group_col: str, base_diff: int, ratio_diff: float, freq: Optional[str]) -> None:
        if group_col == 'month' and 'month' not in self.bank.columns:
            self.bank['month'] = self.bank['date'].dt.to_period('M')
            self.journal['month'] = self.journal['date'].dt.to_period('M')

        keys = set(self.bank[group_col].unique()) & set(
            self.journal[group_col].unique()
        )
        for key in sorted(keys):
            bank_indexes = [
                int(index)
                for index in self.bank[self.bank[group_col] == key].index
            ]
            journal_indexes = [
                int(index)
                for index in self.journal[self.journal[group_col] == key].index
            ]
            if not bank_indexes or not journal_indexes:
                continue
            if len(bank_indexes) == 1 and len(journal_indexes) == 1:
                continue
            bank_amounts = self.bank.loc[
                bank_indexes,
                "amount_decimal",
            ].tolist()
            journal_amounts = self.journal.loc[
                journal_indexes,
                "amount_decimal",
            ].tolist()
            if not self._total_structure_matches(bank_amounts, journal_amounts):
                continue
            metrics = build_group_metrics(bank_amounts, journal_amounts)
            if metrics.group_amount_li == 0:
                continue
            match_type = (
                "monthly_total" if group_col == "month" else "daily_total"
            )
            match_stage = "月核销" if group_col == "month" else "日核销"
            group_evidence = business_evidence(
                [self._business_rows["bank"][index] for index in bank_indexes],
                [self._business_rows["journal"][index] for index in journal_indexes],
            )
            multi_without_hard_boundary = bool(
                len(bank_indexes) > 1
                and len(journal_indexes) > 1
                and not group_evidence["shared_business_id"]
            )
            resolves_full_group = bool(
                not group_evidence["business_conflicts"]
                and not multi_without_hard_boundary
                and self._has_strong_cross_side_evidence(group_evidence)
            )
            self._add_candidate(
                bank_indexes,
                journal_indexes,
                match_type,
                match_stage,
                is_rule_matched=True,
                group_key=str(key),
                resolves_full_group=resolves_full_group,
                complete_business_group=bool(resolves_full_group),
                represents_full_observed_group=bool(
                    group_evidence["business_strength"] > 0
                ),
                total_only_without_boundary=not resolves_full_group,
                batch_boundary_uncertain=bool(
                    multi_without_hard_boundary
                    and group_evidence["business_strength"] > 0
                ),
                bank_distribution=self._get_bucket_dist(bank_amounts),
                journal_distribution=self._get_bucket_dist(journal_amounts),
            )
        self._commit_if_standalone()

    def match_cross_month_total(self) -> None:
        """生成相邻月份边界内的多对多候选，由重要性规则自动分流。"""
        if self.bank.empty or self.journal.empty:
            return
        bank_periods = self.bank["date"].dt.to_period("M")
        journal_periods = self.journal["date"].dt.to_period("M")
        for bank_period, bank_group in self.bank.groupby(bank_periods, sort=True):
            for journal_period, journal_group in self.journal.groupby(
                journal_periods,
                sort=True,
            ):
                bank_month_number = bank_period.year * 12 + bank_period.month
                journal_month_number = (
                    journal_period.year * 12 + journal_period.month
                )
                if abs(bank_month_number - journal_month_number) != 1:
                    continue
                if len(bank_group) <= 1 or len(journal_group) <= 1:
                    continue
                all_dates = [
                    pd.Timestamp(value)
                    for value in (
                        bank_group["date"].tolist()
                        + journal_group["date"].tolist()
                    )
                ]
                if (max(all_dates) - min(all_dates)).days > self.config.tolerance_days:
                    continue
                bank_amounts = bank_group["amount_decimal"].tolist()
                journal_amounts = journal_group["amount_decimal"].tolist()
                if not self._total_structure_matches(
                    bank_amounts,
                    journal_amounts,
                ):
                    continue
                bank_indexes = [int(index) for index in bank_group.index]
                journal_indexes = [int(index) for index in journal_group.index]
                group_evidence = business_evidence(
                    [self._business_rows["bank"][index] for index in bank_indexes],
                    [self._business_rows["journal"][index] for index in journal_indexes],
                )
                resolves_full_group = bool(
                    not group_evidence["business_conflicts"]
                    and (
                        group_evidence["shared_business_id"]
                        or group_evidence["salary_group"]
                        or group_evidence["batch_category"]
                    )
                )
                self._add_candidate(
                    bank_indexes,
                    journal_indexes,
                    "cross_month_total",
                    "跨月核销",
                    is_cross_month_many_to_many=True,
                    is_rule_matched=True,
                    bank_month=str(bank_period),
                    journal_month=str(journal_period),
                    resolves_full_group=resolves_full_group,
                    complete_business_group=resolves_full_group,
                    represents_full_observed_group=bool(
                        group_evidence["business_strength"] > 0
                    ),
                    total_only_without_boundary=not resolves_full_group,
                    batch_boundary_uncertain=bool(
                        not resolves_full_group
                        and group_evidence["business_strength"] > 0
                    ),
                    _precomputed_business_evidence=group_evidence,
                )
        self._commit_if_standalone()

    def _get_bucket_dist(self, amounts):
        return list(bucket_distribution(list(amounts)))
