from decimal import Decimal
import importlib

import data_structures as ds
from precision_engine import PrecisionEngine
from validate import validate_config_params


def _valid_config_kwargs(**overrides):
    values = {
        "tolerance_days": "31",
        "dfs_window": "31",
        "dfs_depth": "30",
        "greedy_attempts": "3",
        "random_seed": "0",
        "similarity_threshold": "0.5",
        "similarity_high": "0.7",
        "max_candidates": "30",
        "memory_limit": "6.0",
        "bank_skip": "0",
        "journal_skip": "0",
        "performance_materiality": "100000",
        "clearly_trivial_threshold": "5000",
        "auto_confirm_score": "70",
    }
    values.update(overrides)
    return values


def test_审计策略使用已确认默认值():
    config = ds.MatcherConfig()

    assert config.performance_materiality == Decimal("100000.00")
    assert config.clearly_trivial_threshold == Decimal("5000.00")
    assert config.auto_confirm_score == 70
    assert config.batch_min_count == 10
    assert config.allow_mixed_sign is False


def test_最终处理状态使用三个正式中文名称():
    assert ds.ProcessingStatus.AUTO_CONFIRMED.value == "自动确认"
    assert ds.ProcessingStatus.PENDING_REVIEW.value == "待人工复核"
    assert ds.ProcessingStatus.NO_CANDIDATE.value == "未找到候选"


def test_差异池使用四个互不抵销的正式分类():
    assert {item.value for item in ds.DifferencePoolType} == {
        "银行已记公司未记-收入",
        "银行已记公司未记-支出",
        "公司已记银行未记-收入",
        "公司已记银行未记-支出",
    }


def test_综合可信度限制在零到一百():
    scores = ds.ScoreBreakdown(amount=40, date=15, text=30, structure=20)

    assert scores.total == 100


def test_实际执行重要性水平不能为负数():
    valid, message = validate_config_params(
        **_valid_config_kwargs(performance_materiality="-1")
    )

    assert valid is False
    assert "实际执行重要性水平" in message


def test_自动确认最低综合可信度不能超过一百():
    valid, message = validate_config_params(
        **_valid_config_kwargs(auto_confirm_score="101")
    )

    assert valid is False
    assert "自动确认最低综合可信度" in message


def _candidate(
    *,
    group_amount: str,
    total_diff: str,
    score: int,
    cross_month: bool = False,
):
    metrics = ds.GroupMetrics(
        bank_gross_li=PrecisionEngine.to_integer_li(group_amount),
        journal_gross_li=PrecisionEngine.to_integer_li(group_amount),
        group_amount_li=PrecisionEngine.to_integer_li(group_amount),
        bank_income_li=0,
        journal_income_li=0,
        bank_expense_li=0,
        journal_expense_li=0,
        income_diff_li=PrecisionEngine.to_integer_li(total_diff),
        expense_diff_li=0,
        total_diff_li=PrecisionEngine.to_integer_li(total_diff),
    )
    return ds.MatchCandidate(
        candidate_id="C1",
        bank_idxs=(0,),
        journal_idxs=(0,),
        match_type="cross_month_total" if cross_month else "tolerance_date",
        match_stage="测试",
        metrics=metrics,
        scores=ds.ScoreBreakdown(
            amount=min(score, 40),
            date=min(max(score - 40, 0), 15),
            text=min(max(score - 55, 0), 30),
            structure=min(max(score - 85, 0), 15),
        ),
        is_cross_month_many_to_many=cross_month,
    )


def _policy():
    return importlib.import_module("matching_policy")


def test_组金额取双方绝对金额合计较大值且收入支出不抵销():
    metrics = _policy().build_group_metrics(
        [
            PrecisionEngine.to_integer_li("120000"),
            PrecisionEngine.to_integer_li("-20000"),
        ],
        [
            PrecisionEngine.to_integer_li("100000"),
            PrecisionEngine.to_integer_li("-10000"),
        ],
    )

    assert metrics.bank_gross_li == 1_400_000_000
    assert metrics.journal_gross_li == 1_100_000_000
    assert metrics.group_amount_li == 1_400_000_000
    assert metrics.income_diff_li == 200_000_000
    assert metrics.expense_diff_li == 100_000_000
    assert metrics.total_diff_li == 300_000_000


def test_大额但差异很小仍然待人工复核():
    candidate = _candidate(group_amount="100000.01", total_diff="1", score=100)

    status, reason = _policy().route_candidate(candidate, ds.MatcherConfig())

    assert status is ds.ProcessingStatus.PENDING_REVIEW
    assert "实际执行重要性水平" in reason


def test_组金额刚好等于实际执行重要性水平不算超过():
    candidate = _candidate(group_amount="100000", total_diff="0", score=100)

    status, _ = _policy().route_candidate(candidate, ds.MatcherConfig())

    assert status is ds.ProcessingStatus.AUTO_CONFIRMED


def test_明显微小错报不受低可信度限制自动处理():
    candidate = _candidate(group_amount="10000", total_diff="5000", score=0)

    status, reason = _policy().route_candidate(candidate, ds.MatcherConfig())

    assert status is ds.ProcessingStatus.AUTO_CONFIRMED
    assert "明显微小错报" in reason


def test_非明显微小错报按七十分门槛分流():
    at_threshold = _candidate(group_amount="10000", total_diff="6000", score=70)
    below_threshold = _candidate(group_amount="10000", total_diff="6000", score=69)

    assert _policy().route_candidate(at_threshold, ds.MatcherConfig())[0] is ds.ProcessingStatus.AUTO_CONFIRMED
    assert _policy().route_candidate(below_threshold, ds.MatcherConfig())[0] is ds.ProcessingStatus.PENDING_REVIEW


def test_跨月多对多即使满分也必须待人工复核():
    candidate = _candidate(
        group_amount="1000",
        total_diff="0",
        score=100,
        cross_month=True,
    )

    status, reason = _policy().route_candidate(candidate, ds.MatcherConfig())

    assert status is ds.ProcessingStatus.PENDING_REVIEW
    assert "跨月多对多" in reason


def test_金额分桶必须按元边界换算为整数厘():
    distribution = _policy().bucket_distribution(
        [
            PrecisionEngine.to_integer_li("50"),
            PrecisionEngine.to_integer_li("100"),
            PrecisionEngine.to_integer_li("1000"),
            PrecisionEngine.to_integer_li("10000"),
            PrecisionEngine.to_integer_li("100000"),
        ]
    )

    assert distribution == (1, 1, 1, 1, 1)


def test_文本评分逐字段比较摘要业务说明和对方户名():
    evidence = _policy().score_text_fields(
        {
            "摘要": "收到成都星河科技有限公司款项",
            "业务说明": "软件服务费",
            "对方户名": "成都星河科技有限公司",
        },
        {
            "摘要": "星河科技软件服务费",
            "业务说明": "收到客户软件服务款",
            "客户名称": "成都星河科技有限公司",
        },
    )

    assert evidence.local_score >= 70
    assert any("对方户名" in item for item in evidence.supporting_fields)
    assert all(
        "银行." in label and "日记账." in label
        for label in evidence.field_scores
    )


def test_文本评分不会把没有字段名的文本粗暴拼接():
    evidence = _policy().score_text_fields(
        {"摘要": "甲乙", "业务说明": "丙丁"},
        {"摘要": "乙丙"},
    )

    assert evidence.local_score < 70
    assert len(evidence.field_scores) == 2


def test_大模型只能补充文本证据且文本项最高三十分():
    evidence = _policy().score_text_fields(
        {"摘要": "苹果采购"},
        {"摘要": "员工差旅"},
        llm_semantic_score=100,
    )

    assert evidence.model_score == 100
    assert evidence.combined_score == 60
    assert evidence.score == 18
    assert evidence.score <= 30


def test_统一评分中金额完全一致得四十分同日得十五分():
    candidate = _candidate(group_amount="10000", total_diff="0", score=0)
    candidate.date_span_days = 0

    scores = _policy().score_candidate(
        candidate,
        ds.MatcherConfig(),
        {"摘要": "销售货款"},
        {"摘要": "销售货款"},
    )

    assert scores.amount == 40
    assert scores.date == 15
    assert candidate.scores == scores
    assert candidate.text_evidence.score == scores.text


def test_日期刚好达到容差上限时日期得分为零():
    candidate = _candidate(group_amount="10000", total_diff="0", score=0)
    candidate.date_span_days = 31

    scores = _policy().score_candidate(
        candidate,
        ds.MatcherConfig(tolerance_days=31),
        {},
        {},
    )

    assert scores.date == 0


def test_歧义候选缺少文本证据时总分不能达到自动确认线():
    candidate = _candidate(group_amount="10000", total_diff="0", score=0)
    candidate.is_ambiguous = True
    candidate.date_span_days = 0

    scores = _policy().score_candidate(
        candidate,
        ds.MatcherConfig(),
        {},
        {},
    )

    assert scores.text == 0
    assert scores.total <= 69


def test_关键账户字段互相矛盾时明确记录冲突():
    evidence = _policy().score_text_fields(
        {"对方户名": "甲公司", "摘要": "服务费"},
        {"对方户名": "乙公司", "摘要": "服务费"},
    )

    assert "对方户名" in evidence.conflicting_fields


def test_敏感账户字段只生成相同不同信号不包含原值():
    signals = _policy().build_sensitive_field_signals(
        {
            "账号": "622233334444",
            "对方户名": "甲公司",
            "卡号": "88889999",
        },
        {
            "帐号": "622233334444",
            "客户名称": "乙公司",
            "卡号": "77776666",
        },
    )

    assert set(signals) == {
        "账号:相同",
        "卡号:不同",
        "对方户名:不同",
    }
    joined = "|".join(signals)
    assert "622233334444" not in joined
    assert "甲公司" not in joined
    assert "乙公司" not in joined
