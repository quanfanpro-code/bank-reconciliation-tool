import json
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import pytest

from llm_assistant import (
    CandidateSemanticRequest,
    LLMConfig,
    LLMAssistant,
    SemanticCandidate,
    redact_sensitive_text,
    sanitize_url,
)


def _decision_data(candidate_id="C1", score=88):
    return {
        "selected_candidate_id": candidate_id,
        "semantic_score": score,
        "reason": "摘要与业务说明一致",
        "supporting_evidence": ["设备款", "购买设备"],
        "conflicting_evidence": [],
        "uncertainty": "低",
        "suggested_status": "自动确认",
    }


def _responses_result(candidate_id="C1", score=88, text=None):
    content = text if text is not None else json.dumps(
        _decision_data(candidate_id, score),
        ensure_ascii=False,
    )
    return {
        "output": [{"content": [{"type": "output_text", "text": content}]}],
        "usage": {"input_tokens": 30, "output_tokens": 20},
    }


def _chat_result(candidate_id="C1", score=88, text=None):
    content = text if text is not None else json.dumps(
        _decision_data(candidate_id, score),
        ensure_ascii=False,
    )
    return {
        "choices": [{"message": {"content": content}}],
        "usage": {"prompt_tokens": 30, "completion_tokens": 20},
    }


class _MockHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def _serve(self):
        content_length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(content_length) if content_length else b""
        body = json.loads(raw_body.decode("utf-8")) if raw_body else None
        self.server.requests.append(
            {
                "method": self.command,
                "path": self.path,
                "body": body,
                "authorization": self.headers.get("Authorization", ""),
            }
        )
        queued = self.server.routes.get(self.path, [])
        if not queued:
            status, payload, delay = 404, {"error": "not found"}, 0
        else:
            status, payload, delay = queued.pop(0)
        if delay:
            time.sleep(delay)
        response = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        try:
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(response)))
            self.end_headers()
            self.wfile.write(response)
        except (BrokenPipeError, ConnectionResetError):
            pass

    do_GET = _serve
    do_POST = _serve


@pytest.fixture
def mock_server():
    server = ThreadingHTTPServer(("127.0.0.1", 0), _MockHandler)
    server.requests = []
    server.routes = {}
    server.url = f"http://127.0.0.1:{server.server_port}"

    def enqueue(path, status, payload, delay=0):
        server.routes.setdefault(path, []).append((status, payload, delay))

    server.enqueue = enqueue
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield server
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=2)


def _sample_request():
    return CandidateSemanticRequest(
        request_id="R1",
        candidates=(
            SemanticCandidate(
                candidate_id="C1",
                bank_date="2026-01-05",
                journal_date="2026-01-06",
                bank_amount=1000.0,
                journal_amount=1000.0,
                bank_fields={
                    "摘要": "设备款",
                    "业务说明": "购置设备",
                    "账号": "完整账号6222",
                    "对方户名": "甲公司",
                },
                journal_fields={
                    "业务说明": "购买设备",
                    "凭证号": "记-001",
                    "对方户名": "甲公司",
                },
                local_signals=(
                    "账号:不同",
                    "对方户名:相同",
                ),
            ),
        ),
        source_file_paths=(r"C:\敏感\银行.xlsx", r"C:\敏感\日记账.xlsx"),
    )


def test_在线请求只发送必要候选字段且密钥只进请求头(mock_server):
    mock_server.enqueue("/v1/responses", 200, _responses_result())
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            mode="online",
            protocol="responses",
            base_url=mock_server.url + "/v1/",
            model="mock-model",
            api_key="secret-key",
        )
    )

    decision = assistant.evaluate_candidates(_sample_request())
    request = mock_server.requests[-1]
    body = json.dumps(request["body"], ensure_ascii=False)

    assert request["authorization"] == "Bearer secret-key"
    assert "secret-key" not in body
    assert "完整账号6222" not in body
    assert "甲公司" not in body
    assert "凭证号" not in body
    assert "C:\\\\" not in body
    assert "账号:不同" in body
    assert "对方户名:相同" in body
    assert decision.selected_candidate_id == "C1"
    assert decision.fallback_used is False
    assert decision.candidate_ids == ("C1",)
    assert set(decision.sent_fields) == {
        "日期",
        "金额",
        "摘要",
        "业务说明",
        "本机敏感字段一致性信号",
    }
    assert decision.protocol == "responses"


def test_本地模式可以发送用户明确选择的辅助字段(mock_server):
    mock_server.enqueue("/v1/chat/completions", 200, _chat_result())
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            mode="local",
            protocol="chat_completions",
            base_url=mock_server.url + "/v1",
            model="local-model",
            local_fields=("摘要", "业务说明", "对方户名"),
        )
    )

    decision = assistant.evaluate_candidates(_sample_request())
    body = json.dumps(mock_server.requests[-1]["body"], ensure_ascii=False)

    assert "对方户名" in body
    assert "甲公司" in body
    assert "完整账号6222" not in body
    assert decision.selected_candidate_id == "C1"


def test_本地模式未选择文字字段时不发送任何文字原值(mock_server):
    mock_server.enqueue("/v1/chat/completions", 200, _chat_result())
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            mode="local",
            protocol="chat_completions",
            base_url=mock_server.url + "/v1",
            model="local-model",
            local_fields=(),
        )
    )

    decision = assistant.evaluate_candidates(_sample_request())
    body = json.dumps(mock_server.requests[-1]["body"], ensure_ascii=False)

    assert "设备款" not in body
    assert "购置设备" not in body
    assert "甲公司" not in body
    assert set(decision.sent_fields) == {
        "日期",
        "金额",
        "本机敏感字段一致性信号",
    }


def test_可以读取模型列表和测试连接(mock_server):
    mock_server.enqueue(
        "/v1/models",
        200,
        {"data": [{"id": "模型乙"}, {"id": "模型甲"}]},
    )
    mock_server.enqueue(
        "/v1/models",
        200,
        {"data": [{"id": "模型甲"}]},
    )
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            base_url=mock_server.url + "/v1",
            api_key="key",
        )
    )

    assert assistant.list_models() == ["模型乙", "模型甲"]
    assert assistant.test_connection()[0] is True


def test_自动协议在响应端点明确不支持时退回聊天端点(mock_server):
    mock_server.enqueue("/v1/responses", 404, {"error": "unsupported"})
    mock_server.enqueue("/v1/chat/completions", 200, _chat_result(score=91))
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            protocol="auto",
            base_url=mock_server.url + "/v1",
            model="mock-model",
        )
    )

    decision = assistant.evaluate_candidates(_sample_request())

    assert [item["path"] for item in mock_server.requests] == [
        "/v1/responses",
        "/v1/chat/completions",
    ]
    assert decision.semantic_score == 91
    assert decision.fallback_used is False


def test_格式错误只修复重试一次(mock_server):
    mock_server.enqueue(
        "/v1/responses",
        200,
        _responses_result(text="这不是JSON"),
    )
    mock_server.enqueue("/v1/responses", 200, _responses_result(score=76))
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            protocol="responses",
            base_url=mock_server.url + "/v1",
            model="mock-model",
        )
    )

    decision = assistant.evaluate_candidates(_sample_request())

    assert len(mock_server.requests) == 2
    assert decision.semantic_score == 76
    assert decision.fallback_used is False


def test_候选编号越界两次后安全降级(mock_server):
    mock_server.enqueue(
        "/v1/responses",
        200,
        _responses_result(candidate_id="不存在"),
    )
    mock_server.enqueue(
        "/v1/responses",
        200,
        _responses_result(candidate_id="仍不存在"),
    )
    assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            protocol="responses",
            base_url=mock_server.url + "/v1",
            model="mock-model",
        )
    )

    decision = assistant.evaluate_candidates(_sample_request())

    assert decision.fallback_used is True
    assert decision.selected_candidate_id == ""
    assert "候选" in decision.error
    assert len(mock_server.requests) == 2


def test_超时和认证失败不会中断核对(mock_server):
    mock_server.enqueue(
        "/v1/responses",
        200,
        _responses_result(),
        delay=0.2,
    )
    timeout_assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            protocol="responses",
            base_url=mock_server.url + "/v1",
            model="mock-model",
            timeout_seconds=0.05,
        )
    )
    timeout_decision = timeout_assistant.evaluate_candidates(_sample_request())

    mock_server.enqueue("/v1/responses", 401, {"error": "bad secret-key"})
    auth_assistant = LLMAssistant(
        LLMConfig(
            enabled=True,
            protocol="responses",
            base_url=mock_server.url + "/v1",
            model="mock-model",
            api_key="secret-key",
        )
    )
    auth_decision = auth_assistant.evaluate_candidates(_sample_request())

    assert timeout_decision.fallback_used is True
    assert "超时" in timeout_decision.error
    assert auth_decision.fallback_used is True
    assert "secret-key" not in auth_decision.error


def test_URL与错误文本脱敏():
    assert sanitize_url(
        "https://user:secret@example.com/v1?api_key=secret#part"
    ) == "https://example.com/v1"
    cleaned = redact_sensitive_text(
        r"文件 C:\客户\银行.xlsx，账号622233334444，密钥 secret-key",
        secrets=("secret-key",),
    )

    assert "C:\\客户" not in cleaned
    assert "622233334444" not in cleaned
    assert "secret-key" not in cleaned
