# ==========================================
# 核心匹配器模块 (Matcher Module)
# ==========================================
# 包含：
#   - 算法函数（DFS、折半枚举、贪心等）
#   - 并行 DFS 处理函数
#   - 核心匹配引擎（Matcher 类）

import gc
import os
import re
import random
import bisect
import secrets
from dataclasses import replace
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
    DifferencePoolResult, LLMDecisionRecord, MatcherConfig,
    MatchCandidate,
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
    
    MAX_SUBSETS = 100000
    left_subsets = {}
    subset_count = [0]
    
    def dfs_left(start_idx, current_sum, path_idxs, path_dates):
        if len(path_idxs) > max_depth:
            return
        if subset_count[0] > MAX_SUBSETS:
            return
        
        key = current_sum
        if key not in left_subsets:
            left_subsets[key] = []
        left_subsets[key].append((list(path_idxs), list(path_dates)))
        subset_count[0] += 1
        
        if subset_count[0] > MAX_SUBSETS:
            return
        
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
    
    if subset_count[0] > MAX_SUBSETS:
        return None
    
    all_solutions = []
    
    def dfs_right(start_idx, current_sum, path_idxs, path_dates):
        if len(path_idxs) > max_depth:
            return
        
        needed = target - current_sum
        if needed in left_subsets:
            for left_idxs, left_dts in left_subsets[needed]:
                combined_idxs = left_idxs + path_idxs
                combined_dates = left_dts + path_dates
                
                if len(combined_idxs) > max_depth:
                    continue
                
                all_solutions.append((combined_idxs, combined_dates))
        
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
    
    if not all_solutions:
        return None
    
    all_solutions.sort(key=lambda x: len(x[0]))
    
    best_idxs, best_dates = all_solutions[0]
    
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
    max_subsets = 100000

    def enumerate_half(
        half_values: List[int],
        half_dates: List[pd.Timestamp],
        half_indices: List[int],
    ) -> Optional[List[Dict[int, Tuple[Tuple[int, ...], Tuple[pd.Timestamp, ...]]]]]:
        by_count: List[
            Dict[int, Tuple[Tuple[int, ...], Tuple[pd.Timestamp, ...]]]
        ] = [dict() for _ in range(min(max_depth, len(half_values)) + 1)]
        subset_count = 0

        def walk(
            position: int,
            current_sum: int,
            path_indices: List[int],
            path_dates: List[pd.Timestamp],
        ) -> None:
            nonlocal subset_count
            if subset_count >= max_subsets:
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
                subset_count += 1

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
        if subset_count >= max_subsets:
            return None
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
    if left_by_count is None or right_by_count is None:
        return None

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
        for left_sum, (left_indices, left_dates) in left_sum_map.items():
            minimum_right_count = max(0, 2 - left_count)
            maximum_right_count = min(
                len(right_by_count) - 1,
                max_depth - left_count,
            )
            for right_count in range(
                minimum_right_count,
                maximum_right_count + 1,
            ):
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


def _process_single_source(args: Tuple) -> Optional[Tuple[int, List[int], str]]:
    """处理单个 source 的 DFS 匹配"""
    source_idx, source_date, target_val, targets_data, config = args
    target_val = int(target_val)
    
    # 解包配置
    max_depth = config.max_dfs_depth
    allow_mixed_sign = config.allow_mixed_sign
    max_candidates = config.max_candidates
    date_window = config.dfs_date_window
    allow_zero_match = config.allow_zero_match
    allow_greedy_fallback = config.allow_greedy_fallback
    greedy_attempts = config.greedy_attempts
    random_seed = config.random_seed

    tgt_view_dict = targets_data['view_dict']

    window_indices: List[int] = []
    window_amounts: List[int] = []
    window_dates: List[Any] = []

    for item in tgt_view_dict:
        if item['matched']: continue
        
        amount_decimal = int(item['amount_decimal'])
        if not allow_mixed_sign and target_val != 0:
            if target_val > 0 and amount_decimal <= 0: continue
            if target_val < 0 and amount_decimal >= 0: continue

        window_indices.append(item['index'])
        window_amounts.append(amount_decimal)
        window_dates.append(item['date'])

    if len(window_indices) < 2:
        return None

    if len(window_indices) > max_candidates:
        indexed = list(zip(window_amounts, window_dates, window_indices))
        indexed.sort(key=lambda x: abs(x[0]), reverse=True)
        indexed = indexed[:max_candidates]
        window_amounts = [item[0] for item in indexed]
        window_dates = [item[1] for item in indexed]
        window_indices = [item[2] for item in indexed]

    if target_val == 0:
        if not allow_zero_match:
            return None
        # 0金额匹配需要特别处理类型
        window_amounts_int = [PrecisionEngine.to_integer_li(a) for a in window_amounts]
        if 0 in window_amounts_int:
            zero_pos = window_amounts_int.index(0)
            return (source_idx, [window_indices[zero_pos]], '中')
        return None

    indexed = list(zip(window_amounts, window_dates, window_indices))
    indexed.sort(key=lambda x: abs(x[0]), reverse=True)
    window_amounts = [item[0] for item in indexed]
    window_dates = [item[1] for item in indexed]
    window_indices = [item[2] for item in indexed]

    max_abs_cand = max(abs(a) for a in window_amounts) if window_amounts else 0
    curr_max_depth = max_depth
    if max_abs_cand > 0:
        curr_max_depth = min(max_depth, 2 + int(abs(target_val) / max_abs_cand))

    trivial_tolerance_li = PrecisionEngine.to_integer_li(
        config.clearly_trivial_threshold
    )
    result_info = _solve_combination(
        list(window_amounts),
        list(window_dates),
        list(window_indices),
        target_val,
        curr_max_depth,
        allow_mixed_sign,
        date_window,
    )
    if result_info is not None and len(result_info[0]) < 2:
        return None
    if result_info is None and trivial_tolerance_li > 0:
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
        return (source_idx, result_idxs, confidence)

    if allow_greedy_fallback:
        greedy_result = _randomized_greedy(
            list(window_amounts), list(window_dates), list(window_indices),
            target_val, greedy_attempts, random_seed
        )
        if greedy_result and len(greedy_result[0]) >= 2:
            result_idxs, confidence = greedy_result
            return (source_idx, result_idxs, confidence)

    return None


# ==========================================
# 核心匹配引擎
# ==========================================


def select_non_conflicting_candidates(
    candidates: List[MatchCandidate],
) -> List[MatchCandidate]:
    """按统一分数稳定排序，确保任何一笔记录只被一个候选占用。"""
    ordered = sorted(
        candidates,
        key=lambda item: (
            -item.scores.total,
            item.metrics.total_diff_li,
            item.date_span_days,
            len(item.bank_idxs) + len(item.journal_idxs),
            item.bank_idxs,
            item.journal_idxs,
            item.candidate_id,
        ),
    )
    used_bank: Set[int] = set()
    used_journal: Set[int] = set()
    selected: List[MatchCandidate] = []
    for candidate in ordered:
        if used_bank.intersection(candidate.bank_idxs):
            continue
        if used_journal.intersection(candidate.journal_idxs):
            continue
        selected.append(candidate)
        used_bank.update(candidate.bank_idxs)
        used_journal.update(candidate.journal_idxs)
    return selected


class Matcher:
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
                 llm_assistant: Any = None):
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
        self._candidate_ids: Set[str] = set()
        self.whitelist_rules = whitelist_rules if whitelist_rules else self.WHITELIST_RULES
        self.llm_assistant = llm_assistant
        requested_seed = self.config.random_seed
        actual_seed = (
            secrets.randbits(64) if requested_seed == -1 else requested_seed
        )
        self.run_parameters = {
            "requested_random_seed": requested_seed,
            "actual_random_seed": actual_seed,
        }
        self._execution_config = replace(self.config, random_seed=actual_seed)
        self._collecting_candidates = False
        self.stopping = False
        
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
        score_candidate(
            candidate,
            self.config,
            self._merged_text_fields(self.bank, candidate.bank_idxs),
            self._merged_text_fields(self.journal, candidate.journal_idxs),
        )

    def _add_candidate(
        self,
        bank_idxs: List[int] | Tuple[int, ...],
        journal_idxs: List[int] | Tuple[int, ...],
        match_type: str,
        match_stage: str,
        **evidence: Any,
    ) -> Optional[MatchCandidate]:
        """只登记候选，不提前占用银行流水或日记账记录。"""
        bank_tuple = tuple(sorted(int(index) for index in bank_idxs))
        journal_tuple = tuple(sorted(int(index) for index in journal_idxs))
        if not bank_tuple or not journal_tuple:
            return None
        candidate_id = self._candidate_id(match_type, bank_tuple, journal_tuple)
        if candidate_id in self._candidate_ids:
            return None

        bank_amounts = self.bank.loc[list(bank_tuple), "amount_decimal"].tolist()
        journal_amounts = self.journal.loc[list(journal_tuple), "amount_decimal"].tolist()
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
            evidence=dict(evidence),
        )
        self._score_existing_candidate(candidate)
        self.candidates.append(candidate)
        self._candidate_ids.add(candidate_id)
        return candidate

    def _refresh_candidate_ambiguity(self) -> None:
        """候选全部生成后，再识别真正存在多个去向的记录。"""
        bank_options: Dict[int, Set[Tuple[int, ...]]] = {}
        journal_options: Dict[int, Set[Tuple[int, ...]]] = {}
        for candidate in self.candidates:
            for bank_index in candidate.bank_idxs:
                bank_options.setdefault(bank_index, set()).add(candidate.journal_idxs)
            for journal_index in candidate.journal_idxs:
                journal_options.setdefault(journal_index, set()).add(candidate.bank_idxs)

        for candidate in self.candidates:
            if candidate.evidence.get("resolves_full_group", False):
                candidate.is_ambiguous = False
            else:
                candidate.is_ambiguous = any(
                    len(bank_options[index]) > 1 for index in candidate.bank_idxs
                ) or any(
                    len(journal_options[index]) > 1
                    for index in candidate.journal_idxs
                )
            self._score_existing_candidate(candidate)

    def _group_ambiguous_candidates(self) -> List[List[MatchCandidate]]:
        """把共享任一银行或日记账记录的候选归入同一竞争组。"""
        bank_members: Dict[int, Set[int]] = {}
        journal_members: Dict[int, Set[int]] = {}
        for position, candidate in enumerate(self.candidates):
            for bank_index in candidate.bank_idxs:
                bank_members.setdefault(bank_index, set()).add(position)
            for journal_index in candidate.journal_idxs:
                journal_members.setdefault(journal_index, set()).add(position)

        groups: List[List[MatchCandidate]] = []
        visited: Set[int] = set()
        for start in range(len(self.candidates)):
            if start in visited:
                continue
            pending = [start]
            component: Set[int] = set()
            while pending:
                position = pending.pop()
                if position in component:
                    continue
                component.add(position)
                candidate = self.candidates[position]
                neighbours: Set[int] = set()
                for bank_index in candidate.bank_idxs:
                    neighbours.update(bank_members.get(bank_index, set()))
                for journal_index in candidate.journal_idxs:
                    neighbours.update(journal_members.get(journal_index, set()))
                pending.extend(neighbours - component)
            visited.update(component)
            groups.append(
                [
                    self.candidates[position]
                    for position in sorted(
                        component,
                        key=lambda item: self.candidates[item].candidate_id,
                    )
                ]
            )
        return groups

    @staticmethod
    def _candidate_sort_key(candidate: MatchCandidate) -> Tuple[Any, ...]:
        return (
            -candidate.scores.total,
            candidate.metrics.total_diff_li,
            candidate.date_span_days,
            len(candidate.bank_idxs) + len(candidate.journal_idxs),
            candidate.bank_idxs,
            candidate.journal_idxs,
            candidate.candidate_id,
        )

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
        self._refresh_candidate_ambiguity()
        self._apply_llm_assistance()
        self.selected_candidates = select_non_conflicting_candidates(self.candidates)
        for sequence, candidate in enumerate(self.selected_candidates, start=1):
            candidate.final_match_id = f"M{sequence:06d}"
            status, reason = route_candidate(candidate, self.config)
            candidate.processing_status = status
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

            self.matches.append(
                {
                    "id": candidate.final_match_id,
                    "type": candidate.match_type,
                    "confidence": confidence,
                    "confidence_score": candidate.scores.total,
                    "processing_status": status.value,
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
                "included_in_pool_review",
                False,
            ):
                continue
            for frame, indices in (
                (self.bank, list(candidate.bank_idxs)),
                (self.journal, list(candidate.journal_idxs)),
            ):
                frame.loc[
                    indices,
                    "processing_status",
                ] = candidate.processing_status.value
            match = match_by_candidate.get(candidate.candidate_id)
            if match is not None:
                match["processing_status"] = (
                    candidate.processing_status.value
                )
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
            bank_indexes = [
                index
                for index in self.bank.index
                if pattern.search(self._rule_search_text(self.bank, int(index)))
                and not self._check_negative_rules(
                    self._rule_search_text(self.bank, int(index))
                )
            ]
            journal_indexes = [
                index
                for index in self.journal.index
                if pattern.search(self._rule_search_text(self.journal, int(index)))
                and not self._check_negative_rules(
                    self._rule_search_text(self.journal, int(index))
                )
            ]
            for bank_index in sorted(bank_indexes):
                bank_amount = self.bank.at[bank_index, "amount_decimal"]
                bank_date = pd.Timestamp(self.bank.at[bank_index, "date"])
                possible = []
                for journal_index in sorted(journal_indexes):
                    journal_amount = self.journal.at[journal_index, "amount_decimal"]
                    if bank_amount != journal_amount:
                        continue
                    journal_date = pd.Timestamp(self.journal.at[journal_index, "date"])
                    date_difference = abs((bank_date - journal_date).days)
                    if date_difference > max_offset:
                        continue
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
                for _, _, journal_index in sorted(possible)[
                    : self.config.max_candidates
                ]:
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
            for name, func in candidate_steps:
                self._log(f"开始: {name}...")
                if self.stopping:
                    self._log("任务已停止")
                    break
                func()
                _gc_cleanup(f"{name}后", self._log)
        finally:
            self._collecting_candidates = False

        if not self.stopping:
            self._commit_selected_candidates()
            
        return self.matches

    def match_exact_1to1(self) -> None:
        """为同日同金额记录生成一对一候选，不在本阶段抢占记录。"""
        b_groups = self.bank.groupby(["date", "amount_decimal"]).groups
        j_groups = self.journal.groupby(["date", "amount_decimal"]).groups
        common_keys = set(b_groups.keys()) & set(j_groups.keys())
        for key in sorted(common_keys, key=lambda item: (item[0], item[1])):
            journal_indexes = sorted(int(index) for index in j_groups[key])
            for bank_index in sorted(int(index) for index in b_groups[key]):
                ranked = []
                for journal_index in journal_indexes:
                    evidence = score_text_fields(
                        self._row_text_fields(self.bank, bank_index),
                        self._row_text_fields(self.journal, journal_index),
                    )
                    ranked.append((-evidence.local_score, journal_index))
                for _, journal_index in sorted(ranked)[
                    : self.config.max_candidates
                ]:
                    self._add_candidate(
                        [bank_index],
                        [journal_index],
                        "exact_1to1",
                        "精确",
                    )

    def match_tolerance(self) -> None:
        """生成日期容差和明显微小金额差异候选。"""
        trivial_li = PrecisionEngine.to_integer_li(
            self.config.clearly_trivial_threshold
        )
        journal_entries = sorted(
            (
                int(row.amount_decimal),
                int(row.Index),
                pd.Timestamp(row.date),
            )
            for row in self.journal[
                ["amount_decimal", "date"]
            ].itertuples()
        )
        journal_amounts = [entry[0] for entry in journal_entries]

        for bank_row in self.bank.sort_index().itertuples():
            bank_index = int(bank_row.Index)
            bank_amount = int(bank_row.amount_decimal)
            bank_date = pd.Timestamp(bank_row.date)
            lower_bound = bisect.bisect_left(
                journal_amounts,
                bank_amount - trivial_li,
            )
            upper_bound = bisect.bisect_right(
                journal_amounts,
                bank_amount + trivial_li,
            )
            ranked = []
            for position in range(lower_bound, upper_bound):
                (
                    journal_amount,
                    journal_index,
                    journal_date,
                ) = journal_entries[position]
                if bank_amount == 0 or journal_amount == 0:
                    if bank_amount != journal_amount:
                        continue
                elif (bank_amount > 0) != (journal_amount > 0):
                    continue
                amount_difference = abs(bank_amount - journal_amount)
                date_difference = abs((bank_date - journal_date).days)
                if amount_difference == 0 and date_difference == 0:
                    continue
                if date_difference > self.config.tolerance_days:
                    continue
                evidence = score_text_fields(
                    self._row_text_fields(self.bank, int(bank_index)),
                    self._row_text_fields(self.journal, int(journal_index)),
                )
                ranked.append(
                    (
                        -evidence.local_score,
                        amount_difference,
                        date_difference,
                        int(journal_index),
                    )
                )
            for _, amount_difference, _, journal_index in sorted(ranked)[
                : self.config.max_candidates
            ]:
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
        if source.empty or target.empty:
            return

        target_by_summary: Dict[str, List[int]] = {}
        for target_index, target_row in target.iterrows():
            summary = normalize_summary(target_row["summary"])
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
            summary = normalize_summary(row["summary"])
            original_row = int(row["original_idx"])
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
        self._dfs_one_to_many('bank', 'journal')
        self._dfs_one_to_many('journal', 'bank')
        self._commit_if_standalone()

    def _dfs_one_to_many(self, source_type: str, target_type: str) -> None:
        if source_type == 'bank':
            sources = self.bank
            targets_df = self.journal
        else:
            sources = self.journal
            targets_df = self.bank
            
        if sources.empty or targets_df.empty: return

        tgt_view = targets_df.reset_index().sort_values('date')
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
            })

        tasks = []
        source_indices = sources.index.tolist()
        date_window = self.config.dfs_date_window
        
        source_dates = sources['date'].to_dict()
        source_amounts = sources['amount_decimal'].to_dict()
        
        for s_idx in source_indices:
            target_val = source_amounts[s_idx]
            s_date = source_dates[s_idx]

            date_min = pd.Timestamp(s_date) - timedelta(days=date_window)
            date_max = pd.Timestamp(s_date) + timedelta(days=date_window)
            date_min_np = np.datetime64(date_min)
            date_max_np = np.datetime64(date_max)
            start_pos = np.searchsorted(tgt_dates, date_min_np, side='left')
            end_pos = np.searchsorted(tgt_dates, date_max_np, side='right')

            window_view_dict = tgt_view_dict[start_pos:end_pos]
            window_dates = tgt_dates[start_pos:end_pos]
            filtered_targets_data = {'dates': window_dates, 'view_dict': window_view_dict}
            tasks.append(
                (
                    s_idx,
                    s_date,
                    target_val,
                    filtered_targets_data,
                    self._execution_config,
                )
            )

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
        results: List[Tuple[int, List[int], str]] = []

        def run_serial_tasks() -> None:
            for task in tasks:
                if self.stopping:
                    break
                try:
                    result = _process_single_source(task)
                    if result:
                        results.append(result)
                except Exception as exc:
                    self.exception_logger.record_exception(
                        f"dfs_{source_type}",
                        int(task[0]),
                        exc,
                    )

        if use_parallel:
            self._log(f"智能组合匹配({source_type}): {len(tasks)} 个任务，使用 {num_workers} 个进程...")
            completed = 0
            try:
                with ProcessPoolExecutor(max_workers=num_workers) as executor:
                    futures = {
                        executor.submit(
                            _process_single_source,
                            task,
                        ): task[0]
                        for task in tasks
                    }
                    for future in as_completed(futures):
                        if self.stopping:
                            executor.shutdown(
                                wait=False,
                                cancel_futures=True,
                            )
                            break
                        task_idx = futures[future]
                        try:
                            res = future.result(timeout=300)
                            if res:
                                results.append(res)
                        except FuturesTimeoutError:
                            self._log("组合任务超时，已跳过")
                            self.exception_logger.record_exception(
                                f"dfs_{source_type}",
                                task_idx,
                                TimeoutError("DFS任务超时"),
                            )
                        except BrokenProcessPool:
                            raise
                        except Exception as exc:
                            self.exception_logger.record_exception(
                                f"dfs_{source_type}",
                                task_idx,
                                exc,
                            )
                        finally:
                            completed += 1
                            self._update_progress(
                                40 + completed / len(tasks) * 20
                            )
            except Exception as exc:
                self._log(
                    f"智能组合匹配({source_type})后台进程不可用，"
                    "已退回单线程继续"
                )
                self.exception_logger.record_exception(
                    f"dfs_{source_type}_parallel",
                    -1,
                    exc,
                )
                if not self.stopping:
                    results.clear()
                    run_serial_tasks()
        else:
            self._log(f"智能组合匹配({source_type}): {len(tasks)} 个任务，单线程处理...")
            run_serial_tasks()

        source_order = {source_idx: order for order, source_idx in enumerate(source_indices)}

        results.sort(
            key=lambda item: (
                self._confidence_sort_key(item[2]),
                len(item[1]),
                source_order.get(item[0], len(source_order)),
                min(item[1]) if item[1] else -1,
            )
        )
        
        for source_idx, matched_idxs, _confidence in results:
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
            self._add_candidate(
                bank_indexes,
                journal_indexes,
                match_type,
                match_stage,
                is_rule_matched=True,
                group_key=str(key),
                resolves_full_group=True,
                bank_distribution=self._get_bucket_dist(bank_amounts),
                journal_distribution=self._get_bucket_dist(journal_amounts),
            )
        self._commit_if_standalone()

    def match_cross_month_total(self) -> None:
        """相邻月份边界内的多对多只生成待复核候选。"""
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
                self._add_candidate(
                    [int(index) for index in bank_group.index],
                    [int(index) for index in journal_group.index],
                    "cross_month_total",
                    "跨月核销",
                    is_cross_month_many_to_many=True,
                    is_rule_matched=True,
                    bank_month=str(bank_period),
                    journal_month=str(journal_period),
                    resolves_full_group=True,
                )
        self._commit_if_standalone()

    def _get_bucket_dist(self, amounts):
        return list(bucket_distribution(list(amounts)))
