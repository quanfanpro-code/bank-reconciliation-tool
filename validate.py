"""
银行流水核对工具 - 参数校验模块

包含高级配置参数和界面输入项的有效性校验函数。
"""


def _validate_int(value: str, name: str, allow_zero: bool = False) -> tuple[bool, str]:
    """校验整数参数。"""
    stripped = value.strip()
    if not stripped:
        return False, f"❌ {name}不能为空"
    if not stripped.lstrip("-").isdigit():
        return False, f"❌ {name}必须为整数，当前值: '{value}'"

    int_value = int(stripped)
    if allow_zero:
        if int_value < 0:
            return False, f"❌ {name}不能为负数，当前值: {int_value}"
    elif int_value <= 0:
        return False, f"❌ {name}必须为正整数，当前值: {int_value}"

    return True, ""


def _validate_float(value: str, name: str, must_positive: bool = True, allow_zero: bool = False) -> tuple[bool, str]:
    """校验浮点参数。"""
    stripped = value.strip()
    if not stripped:
        return False, f"❌ {name}不能为空"

    try:
        float_value = float(stripped)
    except ValueError:
        return False, f"❌ {name}必须为数字，当前值: '{value}'"

    if must_positive:
        if allow_zero:
            if float_value < 0:
                return False, f"❌ {name}不能为负数，当前值: {float_value}"
        elif float_value <= 0:
            return False, f"❌ {name}必须大于 0，当前值: {float_value}"

    return True, ""


def validate_config_params(
    tolerance_days: str,
    dfs_window: str,
    dfs_depth: str,
    greedy_attempts: str,
    *,
    random_seed: str = "0",
    similarity_threshold: str = "0.5",
    similarity_high: str = "0.7",
    max_candidates: str = "30",
    memory_limit: str = "6.0",
    bank_skip: str = "0",
    journal_skip: str = "0",
    performance_materiality: str = "100000",
    clearly_trivial_threshold: str = "5000",
    auto_confirm_score: str = "70",
) -> tuple[bool, str]:
    """
    校验高级参数有效性

    参数:
        tolerance_days: 日期容差（天）
        dfs_window: 组合匹配窗口（天）
        dfs_depth: 组合匹配最大深度
        greedy_attempts: 贪心尝试次数

    返回:
        tuple[bool, str]: (是否有效, 错误信息)
    """
    int_params = [
        (tolerance_days, "日期容差", False),
        (dfs_window, "组合匹配窗口", False),
        (dfs_depth, "组合匹配最大深度", False),
        (greedy_attempts, "贪心尝试次数", False),
        (max_candidates, "最大候选数", False),
        (bank_skip, "银行流水跳过行数", True),
        (journal_skip, "日记账跳过行数", True),
    ]

    for value, name, allow_zero in int_params:
        is_valid, message = _validate_int(value, name, allow_zero=allow_zero)
        if not is_valid:
            return False, message

    # 随机种子：整数且 >= -1（-1 表示每次运行使用不同随机种子）
    seed_stripped = random_seed.strip()
    if not seed_stripped:
        return False, "❌ 随机种子不能为空"
    if not seed_stripped.lstrip("-").isdigit():
        return False, f"❌ 随机种子必须为整数，当前值: '{random_seed}'"
    if int(seed_stripped) < -1:
        return False, f"❌ 随机种子不能小于 -1，当前值: {seed_stripped}"

    # 上限校验（须覆盖 readme/GUI 默认值：窗口 31、深度 30，防止默认值被自己拒绝）
    upper_limits = {
        "日期容差": (tolerance_days, 365),
        "组合匹配窗口": (dfs_window, 90),
        "组合匹配最大深度": (dfs_depth, 100),
        "贪心尝试次数": (greedy_attempts, 1000),
        "最大候选数": (max_candidates, 200),
    }
    for name, (value, limit) in upper_limits.items():
        if int(value.strip()) > limit:
            return False, f"❌ {name}不能超过 {limit}，当前值: {value.strip()}"

    float_params = [
        (similarity_threshold, "相似度阈值", True),
        (similarity_high, "高相似度阈值", True),
        (memory_limit, "内存限制(GB)", False),
    ]
    for value, name, allow_zero in float_params:
        is_valid, message = _validate_float(value, name, allow_zero=allow_zero)
        if not is_valid:
            return False, message

    low_threshold = float(similarity_threshold.strip())
    high_threshold = float(similarity_high.strip())
    if not 0 <= low_threshold <= 1:
        return False, "❌ 相似度阈值必须在 0 到 1 之间"
    if not 0 <= high_threshold <= 1:
        return False, "❌ 高相似度阈值必须在 0 到 1 之间"

    if high_threshold < low_threshold:
        return False, "❌ 高相似度阈值不能小于相似度阈值"

    for value, name in (
        (performance_materiality, "实际执行重要性水平"),
        (clearly_trivial_threshold, "明显微小错报临界值"),
    ):
        is_valid, message = _validate_float(value, name, allow_zero=True)
        if not is_valid:
            return False, message

    is_valid, message = _validate_int(
        auto_confirm_score,
        "自动确认最低综合可信度",
        allow_zero=True,
    )
    if not is_valid:
        return False, message
    if int(auto_confirm_score.strip()) > 100:
        return False, "❌ 自动确认最低综合可信度必须在 0 到 100 之间"

    return (True, "")
