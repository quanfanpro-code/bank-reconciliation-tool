"""
倒排索引加速模块 — 单元测试
"""
import pytest

from matcher import NgramTokenizer, InvertedIndex


def test_2gram分词():
    """2-gram分词结果"""
    tokenizer = NgramTokenizer(n=2)
    result = tokenizer.tokenize("收到货款")
    assert result == ["收到", "到货", "货款"], f"2-gram分词结果不符，实际: {result}"


def test_空字符串分词返回空列表():
    """空字符串分词返回空列表"""
    tokenizer = NgramTokenizer(n=2)
    result = tokenizer.tokenize("")
    assert result == [], f"空字符串分词结果不符，实际: {result}"


def test_单字符分词返回空列表():
    """单字符分词返回空列表"""
    tokenizer = NgramTokenizer(n=2)
    result = tokenizer.tokenize("A")
    assert result == [], f"单字符分词结果不符，实际: {result}"


def test_索引构建和查询():
    """索引构建与查询"""
    index = InvertedIndex(NgramTokenizer(n=2))
    documents = {
        1: "收到货款",
        2: "支付货款",
        3: "转账汇款",
    }
    index.build(documents)

    # 查询"货款"应命中文档1和文档2，不命中文档3
    result1 = index.query("货款")
    assert 1 in result1, "文档1应包含'货款'"
    assert 2 in result1, "文档2应包含'货款'"
    assert 3 not in result1, "文档3不应包含'货款'"

    # get_candidates 至少有1个gram交集
    candidates = index.get_candidates("货款", min_overlap=1)
    assert 1 in candidates, "文档1应是有至少1个gram交集的候选"
    assert 2 in candidates, "文档2应是有至少1个gram交集的候选"


def test_空摘要不崩溃():
    """空摘要/None摘要不崩溃"""
    index = InvertedIndex(NgramTokenizer(n=2))
    documents = {
        1: "",
        2: None,
        3: "正常摘要",
    }
    index.build(documents)

    result = index.query("")
    assert result == set(), f"空查询应返回空集合，实际: {result}"

    result2 = index.query("正常")
    assert 3 in result2, "文档3应被查询到"
