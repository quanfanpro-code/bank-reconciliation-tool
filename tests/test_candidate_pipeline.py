from decimal import Decimal

import pandas as pd

from data_structures import (
    LLMDecisionRecord,
    MatchCandidate,
    MatcherConfig,
    ProcessingStatus,
    ScoreBreakdown,
)
from matcher import Matcher, select_non_conflicting_candidates
from llm_assistant import LLMConfig
from matching_policy import build_group_metrics
from precision_engine import PrecisionEngine


def _std_df(rows):
    """构造匹配器使用的最小标准化数据。"""
    records = []
    for index, (date, amount, fields) in enumerate(rows):
        summary = fields.get("摘要", "")
        records.append(
            {
                "date": pd.Timestamp(date),
                "amount": Decimal(str(amount)),
                "amount_decimal": PrecisionEngine.to_integer_li(amount),
                "summary": summary,
                "aux_text_fields": dict(fields),
                "original_idx": index + 1,
                "original_file_row": index + 2,
                "voucher_no": "",
            }
        )
    return pd.DataFrame(
        records,
        columns=[
            "date",
            "amount",
            "amount_decimal",
            "summary",
            "aux_text_fields",
            "original_idx",
            "original_file_row",
            "voucher_no",
        ],
    )


def _candidate(candidate_id, bank_idxs, journal_idxs, score):
    """构造候选选择测试所需的最小候选。"""
    return MatchCandidate(
        candidate_id=candidate_id,
        bank_idxs=tuple(bank_idxs),
        journal_idxs=tuple(journal_idxs),
        match_type="测试",
        match_stage="测试",
        metrics=build_group_metrics([10000], [10000]),
        scores=ScoreBreakdown(amount=score),
    )


def test_同金额多个候选必须综合日期和文字而非先取最近日期():
    bank = _std_df(
        [("2026-01-10", 100, {"摘要": "甲公司设备款"})]
    )
    journal = _std_df(
        [
            ("2026-01-10", 100, {"摘要": "乙公司房租"}),
            ("2026-01-11", 100, {"摘要": "甲公司设备采购"}),
        ]
    )

    matcher = Matcher(bank, journal, MatcherConfig(tolerance_days=3))
    matcher.run()

    assert matcher.selected_candidates[0].journal_idxs == (1,)


def test_日期容差候选使用金额索引而不是逐行扫描日记账(
    monkeypatch,
):
    bank = _std_df(
        [
            ("2026-01-01", 100, {"摘要": "项目甲"}),
            ("2026-01-02", 200, {"摘要": "项目乙"}),
        ]
    )
    journal = _std_df(
        [
            ("2026-01-01", 100, {"摘要": "项目甲"}),
            ("2026-01-02", 200, {"摘要": "项目乙"}),
        ]
    )
    matcher = Matcher(bank, journal, MatcherConfig())

    def forbid_full_scan():
        raise AssertionError("不应对整张日记账逐行扫描")

    monkeypatch.setattr(matcher.journal, "iterrows", forbid_full_scan)

    matcher.match_tolerance()


def test_一条流水不能被两个候选重复使用():
    selected = select_non_conflicting_candidates(
        [
            _candidate("A", (0,), (0,), 90),
            _candidate("B", (1,), (0,), 80),
        ]
    )

    assert [item.candidate_id for item in selected] == ["A"]


def test_同样输入重复运行的候选和最终编号稳定一致():
    bank = _std_df(
        [
            ("2026-01-02", 100, {"摘要": "甲公司设备款"}),
            ("2026-01-03", 200, {"摘要": "乙公司服务费"}),
        ]
    )
    journal = _std_df(
        [
            ("2026-01-03", 200, {"摘要": "乙公司服务费"}),
            ("2026-01-02", 100, {"摘要": "甲公司设备款"}),
        ]
    )

    first = Matcher(bank, journal, MatcherConfig())
    second = Matcher(bank, journal, MatcherConfig())
    first.run()
    second.run()

    assert [
        (item.candidate_id, item.final_match_id)
        for item in first.selected_candidates
    ] == [
        (item.candidate_id, item.final_match_id)
        for item in second.selected_candidates
    ]


def test_未被选中的候选不能重复占用流水():
    bank = _std_df(
        [
            ("2026-01-10", 100, {"摘要": "甲公司设备款"}),
            ("2026-01-10", 100, {"摘要": "无关备用记录"}),
        ]
    )
    journal = _std_df(
        [("2026-01-10", 100, {"摘要": "甲公司设备款"})]
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()

    assert matcher.bank["matched"].sum() == 1
    assert matcher.journal["matched"].sum() == 1
    assert matcher.bank.loc[1, "matched"] is False or not matcher.bank.loc[1, "matched"]


def test_白名单规则不能绕过金额和收支方向条件():
    bank = _std_df(
        [("2026-01-10", 100, {"摘要": "银行手续费"})]
    )
    journal = _std_df(
        [("2026-01-10", -100, {"摘要": "银行手续费"})]
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()

    assert matcher.selected_candidates == []
    assert matcher.bank["matched"].sum() == 0
    assert matcher.journal["matched"].sum() == 0


def test_银行三笔与日记账五笔同日合计相等形成完整多对多组():
    bank = _std_df(
        [
            ("2026-01-15", amount, {"摘要": "销售回款"})
            for amount in (200, 200, 100)
        ]
    )
    journal = _std_df(
        [
            ("2026-01-15", 100, {"摘要": "销售回款"})
            for _ in range(5)
        ]
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()
    groups = [
        candidate
        for candidate in matcher.selected_candidates
        if candidate.match_type == "daily_total"
    ]

    assert len(groups) == 1
    assert groups[0].bank_idxs == (0, 1, 2)
    assert groups[0].journal_idxs == (0, 1, 2, 3, 4)


def test_批量工资门槛必须由当前日期组单独达到十笔():
    bank_rows = [
        ("2026-01-10", 100, {"摘要": "代发工资"})
        for _ in range(9)
    ] + [
        ("2026-01-11", 100, {"摘要": "代发工资"})
        for _ in range(2)
    ]
    bank = _std_df(bank_rows)
    journal = _std_df(
        [("2026-01-11", 200, {"摘要": "代发工资汇总"})]
    )

    matcher = Matcher(bank, journal, MatcherConfig(batch_min_count=10))
    matcher.run()

    assert not any(
        candidate.match_type == "batch_aggregation"
        for candidate in matcher.selected_candidates
    )


def test_当前日期正好十笔且双方均为工资可形成批量候选():
    bank = _std_df(
        [
            ("2026-01-10", 100, {"摘要": "代发工资"})
            for _ in range(10)
        ]
    )
    journal = _std_df(
        [("2026-01-10", 1000, {"摘要": "本月代发工资汇总"})]
    )

    matcher = Matcher(bank, journal, MatcherConfig(batch_min_count=10))
    matcher.run()

    groups = [
        candidate
        for candidate in matcher.selected_candidates
        if candidate.match_type == "batch_aggregation"
    ]
    assert len(groups) == 1
    assert len(groups[0].bank_idxs) == 10
    assert groups[0].journal_idxs == (0,)


def test_普通一对一跨月但在日期容差内可以自动确认():
    bank = _std_df(
        [("2026-01-31", 100, {"摘要": "甲公司货款"})]
    )
    journal = _std_df(
        [("2026-02-01", 100, {"摘要": "甲公司货款"})]
    )

    matcher = Matcher(bank, journal, MatcherConfig(tolerance_days=3))
    matcher.run()

    assert len(matcher.selected_candidates) == 1
    assert matcher.selected_candidates[0].match_type == "tolerance_date"
    assert (
        matcher.selected_candidates[0].processing_status
        is ProcessingStatus.AUTO_CONFIRMED
    )


def test_跨月多对多即使合计一致也只能待人工复核():
    bank = _std_df(
        [
            ("2026-01-31", 60, {"摘要": "项目回款"}),
            ("2026-01-31", 40, {"摘要": "项目回款"}),
        ]
    )
    journal = _std_df(
        [
            ("2026-02-01", 30, {"摘要": "项目回款"}),
            ("2026-02-01", 70, {"摘要": "项目回款"}),
        ]
    )

    matcher = Matcher(bank, journal, MatcherConfig(tolerance_days=3))
    matcher.run()
    groups = [
        candidate
        for candidate in matcher.selected_candidates
        if candidate.match_type == "cross_month_total"
    ]

    assert len(groups) == 1
    assert groups[0].processing_status is ProcessingStatus.PENDING_REVIEW


def test_一坨做账组合存在微小金额差异时仍可自动形成整组():
    bank = _std_df(
        [("2026-01-15", 50000, {"摘要": "批次货款"})]
    )
    journal = _std_df(
        [
            ("2026-01-15", 20000, {"摘要": "批次货款"}),
            ("2026-01-15", 20000, {"摘要": "批次货款"}),
            ("2026-01-15", 9900, {"摘要": "批次货款"}),
        ]
    )

    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.run()

    assert len(matcher.selected_candidates) == 1
    candidate = matcher.selected_candidates[0]
    assert candidate.match_type == "combination_dfs"
    assert candidate.bank_idxs == (0,)
    assert candidate.journal_idxs == (0, 1, 2)
    assert candidate.metrics.total_diff_li == (
        PrecisionEngine.to_integer_li("100")
    )
    assert (
        candidate.processing_status
        is ProcessingStatus.AUTO_CONFIRMED
    )


def test_明显微小差异组合在精确搜索无结果后继续近似搜索(
    monkeypatch,
):
    bank = _std_df(
        [("2026-01-15", 50000, {"摘要": "批次货款"})]
    )
    journal = _std_df(
        [
            ("2026-01-15", 20000, {"摘要": "批次货款"}),
            ("2026-01-15", 20000, {"摘要": "批次货款"}),
            ("2026-01-15", 9900, {"摘要": "批次货款"}),
        ]
    )

    exact_search_calls = []

    def no_exact_combination(*args, **_kwargs):
        exact_search_calls.append(args)
        return None

    monkeypatch.setattr(
        "matcher._solve_combination",
        no_exact_combination,
    )
    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.match_dfs_combinations()

    assert any(
        candidate.match_type == "combination_dfs"
        for candidate in matcher.candidates
    )
    assert exact_search_calls


def test_小批量组合匹配不会启动后台多进程(monkeypatch):
    bank = _std_df(
        [("2026-01-15", 100, {"摘要": "小批量核销"})]
    )
    journal = _std_df(
        [
            ("2026-01-15", 60, {"摘要": "小批量核销"}),
            ("2026-01-15", 40, {"摘要": "小批量核销"}),
        ]
    )

    def forbid_process_pool(*_args, **_kwargs):
        raise AssertionError("小批量不应启动后台多进程")

    monkeypatch.setattr(
        "matcher.ProcessPoolExecutor",
        forbid_process_pool,
    )
    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.match_dfs_combinations()

    assert any(
        candidate.match_type == "combination_dfs"
        for candidate in matcher.candidates
    )


def test_组合阶段不会把一对一精确匹配重复包装成组合():
    bank = _std_df(
        [("2026-01-15", 100, {"摘要": "唯一精确项"})]
    )
    journal = _std_df(
        [
            ("2026-01-15", 100, {"摘要": "唯一精确项"}),
            ("2026-01-15", 10000, {"摘要": "无关大额项"}),
        ]
    )
    matcher = Matcher(bank, journal, MatcherConfig())

    matcher.match_dfs_combinations()

    assert not any(
        candidate.match_type == "combination_dfs"
        and len(candidate.journal_idxs) == 1
        for candidate in matcher.candidates
    )


def test_存在一对一精确项时不再穷举近似组合(monkeypatch):
    bank = _std_df(
        [("2026-01-15", 100, {"摘要": "唯一精确项"})]
    )
    journal = _std_df(
        [
            ("2026-01-15", 100, {"摘要": "唯一精确项"}),
            ("2026-01-15", 10000, {"摘要": "无关大额项"}),
        ]
    )
    near_search_calls = []

    def record_near_search(*args, **_kwargs):
        near_search_calls.append(args)
        return None

    monkeypatch.setattr(
        "matcher._near_combination_solve",
        record_near_search,
    )
    matcher = Matcher(bank, journal, MatcherConfig())
    matcher.match_dfs_combinations()

    assert near_search_calls == []


def test_后台多进程启动失败时自动退回单线程(monkeypatch):
    bank = _std_df(
        [("2026-01-15", 100, {"摘要": "稳定性核销"})]
    )
    journal = _std_df(
        [
            ("2026-01-15", 60, {"摘要": "稳定性核销"}),
            ("2026-01-15", 40, {"摘要": "稳定性核销"}),
        ]
    )
    logs = []

    def fail_process_pool(*_args, **_kwargs):
        raise OSError("模拟后台进程不可用")

    monkeypatch.setattr(
        "matcher.ProcessPoolExecutor",
        fail_process_pool,
    )
    monkeypatch.setattr("matcher.os.cpu_count", lambda: 1)
    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        logger=logs.append,
    )
    matcher.PARALLEL_MIN_TASKS = 1
    matcher.PARALLEL_TASKS_PER_WORKER = 1
    matcher.match_dfs_combinations()

    assert any("退回单线程" in message for message in logs)
    assert any(
        candidate.match_type == "combination_dfs"
        for candidate in matcher.candidates
    )


def test_随机种子负一会生成并记录本次真实种子():
    empty = _std_df([])
    matcher = Matcher(empty, empty, MatcherConfig(random_seed=-1))

    assert matcher.run_parameters["requested_random_seed"] == -1
    assert matcher.run_parameters["actual_random_seed"] >= 0
    assert matcher.run_parameters["actual_random_seed"] != -1


def test_匹配器金额分桶按元边界而不是把整数厘当元():
    empty = _std_df([])
    matcher = Matcher(empty, empty, MatcherConfig())

    assert matcher._get_bucket_dist(
        [
            PrecisionEngine.to_integer_li("50"),
            PrecisionEngine.to_integer_li("100"),
            PrecisionEngine.to_integer_li("1000"),
            PrecisionEngine.to_integer_li("10000"),
            PrecisionEngine.to_integer_li("100000"),
        ]
    ) == [1, 1, 1, 1, 1]


class _FakeAssistant:
    def __init__(
        self,
        *,
        semantic_score=100,
        suggested_status="自动确认",
        candidate_limit=5,
        invalid_candidate=False,
        fallback=False,
    ):
        self.semantic_score = semantic_score
        self.suggested_status = suggested_status
        self.candidate_limit = candidate_limit
        self.invalid_candidate = invalid_candidate
        self.fallback = fallback
        self.requests = []

    def evaluate_candidates(self, request):
        self.requests.append(request)
        if self.fallback:
            return LLMDecisionRecord(
                request_id=request.request_id,
                fallback_used=True,
                error="请求超时，已使用本地文字评分",
            )
        selected = (
            "候选集外编号"
            if self.invalid_candidate
            else request.candidates[-1].candidate_id
        )
        return LLMDecisionRecord(
            request_id=request.request_id,
            selected_candidate_id=selected,
            semantic_score=self.semantic_score,
            suggested_status=self.suggested_status,
            reason="模型语义判断",
        )


def test_大模型不能把超过实际执行重要性水平的候选改成自动确认():
    assistant = _FakeAssistant()
    bank = _std_df(
        [("2026-01-05", 120000, {"摘要": "甲公司设备款"})]
    )
    journal = _std_df(
        [
            ("2026-01-05", 120000, {"摘要": "甲公司设备采购"}),
            ("2026-01-06", 120000, {"摘要": "甲公司预付款"}),
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()

    assert assistant.requests == []
    assert (
        matcher.selected_candidates[0].processing_status
        is ProcessingStatus.PENDING_REVIEW
    )


def test_大模型超时后使用本地评分且核对继续():
    assistant = _FakeAssistant(fallback=True)
    bank = _std_df(
        [("2026-01-05", 1000, {"摘要": "设备款"})]
    )
    journal = _std_df(
        [
            ("2026-01-05", 1000, {"摘要": "购置设备"}),
            ("2026-01-06", 1000, {"摘要": "材料采购"}),
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    result = matcher.run()

    assert result
    assert matcher.llm_records[0].fallback_used is True
    assert matcher.selected_candidates


def test_本地大模型空字段降级报告只记录日期和金额():
    assistant = _FakeAssistant(fallback=True)
    assistant.config = LLMConfig(
        enabled=True,
        mode="local",
        local_fields=(),
    )
    bank = _std_df(
        [("2026-01-05", 1000, {"摘要": "设备款"})]
    )
    journal = _std_df(
        [
            ("2026-01-05", 1000, {"摘要": "购置设备"}),
            ("2026-01-06", 1000, {"摘要": "材料采购"}),
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()

    assert set(matcher.llm_records[0].sent_fields) == {"日期", "金额"}


def test_唯一且文字证据充分的精确候选不调用大模型():
    assistant = _FakeAssistant()
    bank = _std_df(
        [("2026-01-05", 1000, {"摘要": "甲公司设备采购"})]
    )
    journal = _std_df(
        [("2026-01-05", 1000, {"摘要": "甲公司设备采购"})]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()

    assert assistant.requests == []
    assert matcher.llm_records == []


def test_大模型返回候选集外编号会被忽略并记录降级():
    assistant = _FakeAssistant(invalid_candidate=True)
    bank = _std_df(
        [("2026-01-05", 1000, {"摘要": "设备款"})]
    )
    journal = _std_df(
        [
            ("2026-01-05", 1000, {"摘要": "购置设备"}),
            ("2026-01-06", 1000, {"摘要": "材料采购"}),
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()

    assert matcher.llm_records[0].fallback_used is True
    assert "候选集以外" in matcher.llm_records[0].error
    assert all(
        candidate.text_evidence.model_score is None
        for candidate in matcher.candidates
    )


def test_交给大模型的候选数量受配置上限约束():
    assistant = _FakeAssistant(candidate_limit=2)
    bank = _std_df(
        [("2026-01-05", 1000, {"摘要": "设备款"})]
    )
    journal = _std_df(
        [
            (f"2026-01-{day:02d}", 1000, {"摘要": f"候选{day}"})
            for day in range(5, 9)
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()

    assert len(assistant.requests) == 1
    assert len(assistant.requests[0].candidates) == 2


def test_大模型评分只按四六权重重算文字项():
    assistant = _FakeAssistant(semantic_score=100)
    bank = _std_df(
        [("2026-01-05", 1000, {"摘要": "设备款"})]
    )
    journal = _std_df(
        [
            ("2026-01-05", 1000, {"摘要": "购置设备"}),
            ("2026-01-06", 1000, {"摘要": "材料采购"}),
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()
    model_candidate = next(
        candidate
        for candidate in matcher.candidates
        if candidate.llm_decision is not None
        and not candidate.llm_decision.fallback_used
    )

    expected_combined = round(
        model_candidate.text_evidence.local_score * 0.4 + 100 * 0.6
    )
    assert model_candidate.text_evidence.model_score == 100
    assert model_candidate.text_evidence.combined_score == expected_combined
    assert model_candidate.scores.amount == 40


def test_交给大模型的账户证据只有本机相同不同信号():
    assistant = _FakeAssistant()
    bank = _std_df(
        [
            (
                "2026-01-05",
                1000,
                {"摘要": "设备款", "账号": "622233334444"},
            )
        ]
    )
    journal = _std_df(
        [
            (
                "2026-01-05",
                1000,
                {"摘要": "购置设备", "帐号": "622233334444"},
            ),
            (
                "2026-01-06",
                1000,
                {"摘要": "材料采购", "帐号": "999988887777"},
            ),
        ]
    )

    matcher = Matcher(
        bank,
        journal,
        MatcherConfig(),
        llm_assistant=assistant,
    )
    matcher.run()

    signals = {
        signal
        for candidate in assistant.requests[0].candidates
        for signal in candidate.local_signals
    }
    assert signals == {"账号:相同", "账号:不同"}
    assert "622233334444" not in "|".join(signals)
    assert "999988887777" not in "|".join(signals)
