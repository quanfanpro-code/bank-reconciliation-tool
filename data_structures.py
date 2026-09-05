import threading
from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum
from typing import List, Dict, Any, Optional
import pandas as pd

# ==========================================
# 模块级常量 (Module Constants)
# ==========================================

# Excel日期序列号范围
EXCEL_DATE_MIN = 1
EXCEL_DATE_MAX = 100000  # 支持到2173年（约274年范围）

# DFS置信度阈值
DFS_CONFIDENCE_HIGH_THRESHOLD = 2  # 组合数<=2为高置信度
DFS_CONFIDENCE_MEDIUM_THRESHOLD = 5  # 组合数<=5为中置信度

# 折半枚举算法阈值
MEET_IN_MIDDLE_THRESHOLD = 15  # 候选数超过此值时使用折半枚举算法

# 日期偏移量优先级
DATE_OFFSET_PRIORITIES = [0, 1, 3]  # 优先级: 同日(0) -> 相差1天 -> 相差3天

# ==========================================
# 数据结构 (Data Structures)
# ==========================================


class ProcessingStatus(str, Enum):
    """核对结果的最终处理状态。"""

    AUTO_CONFIRMED = "自动确认"
    GROUP_RECONCILED = "整组勾稽一致"
    AUTO_CLASSIFIED = "自动归集事项"
    FLAGGED = "疑点事项"
    NO_CANDIDATE = "未找到候选"


class RiskLevel(str, Enum):
    """程序自动评定的业务风险等级。"""

    NORMAL = "正常"
    LOW = "低风险"
    MEDIUM = "中风险"
    HIGH = "高风险"
    UNKNOWN = "范围未知"


class DifferencePoolType(str, Enum):
    """明显微小错报的四个独立累计池。"""

    BANK_ONLY_INCOME = "银行已记公司未记-收入"
    BANK_ONLY_EXPENSE = "银行已记公司未记-支出"
    JOURNAL_ONLY_INCOME = "公司已记银行未记-收入"
    JOURNAL_ONLY_EXPENSE = "公司已记银行未记-支出"


@dataclass(frozen=True)
class ScoreBreakdown:
    """综合可信度的四项分数。"""

    amount: int = 0
    date: int = 0
    text: int = 0
    structure: int = 0

    @property
    def total(self) -> int:
        return max(0, min(100, self.amount + self.date + self.text + self.structure))


@dataclass(frozen=True)
class TextEvidence:
    """逐字段文本比较的分数、支持证据和冲突证据。"""

    field_scores: Dict[str, int] = field(default_factory=dict)
    supporting_fields: tuple[str, ...] = ()
    conflicting_fields: tuple[str, ...] = ()
    local_score: int = 0
    model_score: Optional[int] = None
    combined_score: int = 0
    score: int = 0


@dataclass(frozen=True)
class GroupMetrics:
    """匹配组两侧金额和不抵销差异。"""

    bank_gross_li: int
    journal_gross_li: int
    group_amount_li: int
    bank_income_li: int
    journal_income_li: int
    bank_expense_li: int
    journal_expense_li: int
    income_diff_li: int
    expense_diff_li: int
    total_diff_li: int


@dataclass
class LLMDecisionRecord:
    """一次大模型辅助判断的完整可追溯记录。"""

    request_id: str
    candidate_ids: tuple[str, ...] = ()
    sent_fields: tuple[str, ...] = ()
    protocol: str = ""
    selected_candidate_id: str = ""
    semantic_score: int = 0
    reason: str = ""
    supporting_evidence: tuple[str, ...] = ()
    conflicting_evidence: tuple[str, ...] = ()
    uncertainty: str = ""
    suggested_status: str = ""
    provider: str = ""
    model: str = ""
    started_at: str = ""
    duration_ms: int = 0
    usage: Dict[str, int] = field(default_factory=dict)
    fallback_used: bool = False
    error: str = ""
    raw_response: str = ""


@dataclass
class MatchCandidate:
    """匹配算法产生、等待统一评分和分流的候选组。"""

    candidate_id: str
    bank_idxs: tuple[int, ...]
    journal_idxs: tuple[int, ...]
    match_type: str
    match_stage: str
    metrics: GroupMetrics
    bank_dates: tuple[pd.Timestamp, ...] = ()
    journal_dates: tuple[pd.Timestamp, ...] = ()
    date_span_days: int = 0
    scores: ScoreBreakdown = field(default_factory=ScoreBreakdown)
    is_cross_month_many_to_many: bool = False
    is_ambiguous: bool = False
    rule_matched: bool = False
    processing_status: ProcessingStatus = ProcessingStatus.FLAGGED
    risk_level: RiskLevel = RiskLevel.NORMAL
    processing_reason: str = ""
    final_match_id: str = ""
    text_evidence: Optional[TextEvidence] = None
    llm_decision: Optional[LLMDecisionRecord] = None
    evidence: Dict[str, Any] = field(default_factory=dict)
    stable_key: str = ""
    composition_key: str = ""


@dataclass
class DifferenceComponent:
    """一笔候选拆入某个自然月差异池的组成。"""

    month: str
    pool_type: DifferencePoolType
    candidate_id: str
    diff_li: int
    included_in_risk_pool: bool = False


@dataclass
class DifferencePoolResult:
    """一个自然月、一个来源方向的明显微小错报累计池。"""

    pool_id: str
    month: str
    pool_type: DifferencePoolType
    total_diff_li: int = 0
    components: List[DifferenceComponent] = field(default_factory=list)
    exceeds_performance_materiality: bool = False
    processing_status: ProcessingStatus = ProcessingStatus.AUTO_CLASSIFIED
    risk_level: RiskLevel = RiskLevel.NORMAL
    processing_reason: str = ""


@dataclass
class MatcherConfig:
    """匹配器配置类"""
    allow_mixed_sign: bool = False
    dfs_date_window: int = 31
    max_dfs_depth: int = 30
    tolerance_days: int = 31
    allow_zero_match: bool = False
    allow_greedy_fallback: bool = True
    greedy_attempts: int = 3
    memory_limit_gb: float = 6.0
    similarity_threshold: float = 0.5
    similarity_high_threshold: float = 0.7
    max_candidates: int = 30
    combination_node_limit_per_source: int = 100000
    combination_task_timeout_seconds: float = 30.0
    combination_global_time_limit_seconds: float = 300.0
    random_seed: int = 0
    whitelist_rules: Optional[List[Dict[str, Any]]] = None
    performance_materiality: Decimal = Decimal("100000.00")
    clearly_trivial_threshold: Decimal = Decimal("5000.00")
    auto_confirm_score: int = 70
    batch_min_count: int = 10

class WorkerExceptionLogger:
    """Worker异常记录器 - 用于记录并行匹配过程中的异常"""
    
    MAX_EXCEPTIONS = 10
    
    def __init__(self):
        self._exceptions: List[Dict[str, Any]] = []
        self._lock = threading.Lock()
        self._counts_by_stage: Dict[str, int] = {}
    
    def record_exception(self, stage: str, task_idx: int, exception: Exception) -> None:
        with self._lock:
            self._counts_by_stage[stage] = self._counts_by_stage.get(stage, 0) + 1
            if len(self._exceptions) < self.MAX_EXCEPTIONS:
                self._exceptions.append({
                    'stage': stage,
                    'task_idx': task_idx,
                    'exception_type': type(exception).__name__,
                    'message': str(exception)
                })
    
    def get_summary(self) -> dict:
        with self._lock:
            return {
                'total_exceptions': sum(self._counts_by_stage.values()),
                'counts_by_stage': dict(self._counts_by_stage),
                'samples': list(self._exceptions)
            }


@dataclass
class DailyBalance:
    """每日余额记录"""
    date: pd.Timestamp
    income: Decimal
    expense: Decimal
    net: Decimal
    balance: Decimal
    prev_balance: Decimal


@dataclass
class BalanceDiff:
    """余额差异记录"""
    date: pd.Timestamp
    bank_balance: Decimal
    journal_balance: Decimal
    diff: Decimal
    diff_type: str


@dataclass
class InitialBalanceWarning:
    """期初余额警告信息"""
    has_warning: bool = False
    bank_initial: Decimal = Decimal('0')
    journal_initial: Decimal = Decimal('0')
    diff: Decimal = Decimal('0')
    message: str = ""


@dataclass(frozen=True)
class OverallControlResult:
    """匹配前形成的期间、收支和余额总体控制结果。"""

    bank_start_date: Optional[pd.Timestamp] = None
    bank_end_date: Optional[pd.Timestamp] = None
    journal_start_date: Optional[pd.Timestamp] = None
    journal_end_date: Optional[pd.Timestamp] = None
    period_status: str = "未实施"
    bank_income: Decimal = Decimal('0')
    bank_expense: Decimal = Decimal('0')
    bank_net: Decimal = Decimal('0')
    journal_income: Decimal = Decimal('0')
    journal_expense: Decimal = Decimal('0')
    journal_net: Decimal = Decimal('0')
    amount_status: str = "未实施"
    bank_initial_balance: Optional[Decimal] = None
    bank_ending_balance: Optional[Decimal] = None
    bank_expected_ending_balance: Optional[Decimal] = None
    bank_balance_diff: Optional[Decimal] = None
    bank_balance_status: str = "未实施"
    journal_initial_balance: Optional[Decimal] = None
    journal_ending_balance: Optional[Decimal] = None
    journal_expected_ending_balance: Optional[Decimal] = None
    journal_balance_diff: Optional[Decimal] = None
    journal_balance_status: str = "未实施"
    initial_balance_diff: Optional[Decimal] = None
    ending_balance_diff: Optional[Decimal] = None
    continuity_anomalies: tuple[Dict[str, Any], ...] = ()
    scope_limited: bool = False
    reasons: tuple[str, ...] = ()

    @property
    def balance_check_possible(self) -> bool:
        """双方都有余额数据时才可执行双侧余额核对。"""
        return (
            self.bank_balance_status != "未实施"
            and self.journal_balance_status != "未实施"
        )
