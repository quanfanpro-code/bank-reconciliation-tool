"""可选的大模型语义辅助客户端，失败时始终安全退回本地规则。"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import os
import re
import socket
import time
from typing import Any
from urllib import error, request
from urllib.parse import urlsplit, urlunsplit

from data_structures import LLMDecisionRecord


ONLINE_ALLOWED_FIELDS = frozenset(
    {"摘要", "业务说明", "交易用途", "附言", "备注"}
)
REQUIRED_DECISION_KEYS = frozenset(
    {
        "selected_candidate_id",
        "semantic_score",
        "reason",
        "supporting_evidence",
        "conflicting_evidence",
        "uncertainty",
        "suggested_status",
    }
)
ALLOWED_SUGGESTED_STATUSES = frozenset(
    {"自动确认", "待人工复核", "未找到候选", ""}
)


@dataclass(frozen=True)
class LLMConfig:
    """大模型连接和候选限制配置，仅存在于当前程序内存。"""

    enabled: bool = False
    mode: str = "online"
    protocol: str = "auto"
    base_url: str = ""
    model: str = ""
    api_key: str = ""
    timeout_seconds: float = 30
    candidate_limit: int = 5
    local_fields: tuple[str, ...] = ()


@dataclass(frozen=True)
class SemanticCandidate:
    """交给大模型比较的最小候选信息。"""

    candidate_id: str
    bank_date: str
    journal_date: str
    bank_amount: float
    journal_amount: float
    bank_fields: dict[str, str]
    journal_fields: dict[str, str]
    local_signals: tuple[str, ...] = ()


@dataclass(frozen=True)
class CandidateSemanticRequest:
    """一次含糊候选组的语义判断请求。"""

    request_id: str
    candidates: tuple[SemanticCandidate, ...]
    source_file_paths: tuple[str, ...] = ()


class LLMRequestError(RuntimeError):
    """统一承载网络和服务端错误，不向核对主流程扩散。"""

    def __init__(self, message: str, status: int = 0):
        super().__init__(message)
        self.status = status


def sanitize_url(value: str) -> str:
    """移除地址中的账号、密钥参数和片段，只保留服务位置。"""
    try:
        parts = urlsplit(str(value).strip())
    except ValueError:
        return ""
    if not parts.scheme or not parts.hostname:
        return ""
    host = parts.hostname
    if ":" in host and not host.startswith("["):
        host = f"[{host}]"
    if parts.port:
        host = f"{host}:{parts.port}"
    path = parts.path.rstrip("/")
    return urlunsplit((parts.scheme, host, path, "", ""))


def redact_sensitive_text(
    value: object,
    *,
    secrets: tuple[str, ...] = (),
) -> str:
    """清理错误和原始回答中的密钥、文件路径及账户号码。"""
    text = str(value or "")
    for secret in secrets:
        if secret:
            text = text.replace(secret, "***")
    text = re.sub(
        r"[A-Za-z]:\\[^\s,，;；\"']+",
        "[路径已隐藏]",
        text,
    )
    text = re.sub(
        r"(账号|帐号|卡号)\s*[:：]?\s*[0-9A-Za-z-]{4,}",
        r"\1***",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\b\d{8,}\b", "***", text)
    text = re.sub(
        r"Bearer\s+[A-Za-z0-9._~+/=-]+",
        "Bearer ***",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"([?&](?:api[_-]?key|access[_-]?token|auth|authorization)=)"
        r"[^&#\s]+",
        r"\1***",
        text,
        flags=re.IGNORECASE,
    )
    return text


def _filter_fields(
    fields: dict[str, str],
    *,
    mode: str,
    local_fields: tuple[str, ...],
) -> dict[str, str]:
    if mode == "local":
        allowed = set(local_fields)
    else:
        allowed = set(ONLINE_ALLOWED_FIELDS)
    result = {}
    for label, value in fields.items():
        if label not in allowed:
            continue
        cleaned = (
            str(value)
            if mode == "local"
            else redact_sensitive_text(value)
        )
        if cleaned.strip():
            result[str(label)] = cleaned.strip()
    return result


def build_online_payload(
    semantic_request: CandidateSemanticRequest,
    *,
    mode: str = "online",
    local_fields: tuple[str, ...] = (),
    candidate_limit: int | None = None,
) -> dict[str, Any]:
    """只构造候选所需数据，主动舍弃路径、账号和凭证号。"""
    candidates = semantic_request.candidates
    if candidate_limit is not None:
        candidates = candidates[:max(0, candidate_limit)]
    return {
        "request_id": semantic_request.request_id,
        "candidates": [
            {
                "candidate_id": item.candidate_id,
                "bank_date": item.bank_date,
                "journal_date": item.journal_date,
                "bank_amount": item.bank_amount,
                "journal_amount": item.journal_amount,
                "bank_fields": _filter_fields(
                    item.bank_fields,
                    mode=mode,
                    local_fields=local_fields,
                ),
                "journal_fields": _filter_fields(
                    item.journal_fields,
                    mode=mode,
                    local_fields=local_fields,
                ),
                "local_signals": list(item.local_signals),
            }
            for item in candidates
        ],
    }


def _decision_schema() -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "selected_candidate_id": {"type": "string"},
            "semantic_score": {
                "type": "integer",
                "minimum": 0,
                "maximum": 100,
            },
            "reason": {"type": "string"},
            "supporting_evidence": {
                "type": "array",
                "items": {"type": "string"},
            },
            "conflicting_evidence": {
                "type": "array",
                "items": {"type": "string"},
            },
            "uncertainty": {"type": "string"},
            "suggested_status": {
                "type": "string",
                "enum": sorted(ALLOWED_SUGGESTED_STATUSES),
            },
        },
        "required": sorted(REQUIRED_DECISION_KEYS),
    }


class LLMAssistant:
    """使用 OpenAI 兼容协议调用在线服务或本地 LM Studio。"""

    def __init__(self, config: LLMConfig):
        self.config = config
        configured_url = config.base_url
        if not configured_url and config.mode == "local":
            configured_url = "http://localhost:1234/v1"
        self.base_url = sanitize_url(configured_url)
        self.api_key = (
            config.api_key
            or os.environ.get("BANK_RECONCILIATION_API_KEY", "")
            or os.environ.get("OPENAI_API_KEY", "")
        )
        self.last_error = ""

    @property
    def provider_name(self) -> str:
        return "LM Studio" if self.config.mode == "local" else "在线 API"

    def _headers(self) -> dict[str, str]:
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
        }
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        return headers

    def _request_json(
        self,
        method: str,
        path: str,
        payload: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        if not self.base_url:
            raise LLMRequestError("服务地址为空或格式不正确")
        body = (
            json.dumps(payload, ensure_ascii=False).encode("utf-8")
            if payload is not None
            else None
        )
        http_request = request.Request(
            self.base_url + path,
            data=body,
            headers=self._headers(),
            method=method,
        )
        try:
            with request.urlopen(
                http_request,
                timeout=max(0.01, self.config.timeout_seconds),
            ) as response:
                raw = response.read().decode("utf-8", errors="replace")
        except error.HTTPError as exc:
            raw_error = exc.read().decode("utf-8", errors="replace")
            safe_error = redact_sensitive_text(
                raw_error,
                secrets=(self.api_key,),
            )
            raise LLMRequestError(
                f"服务返回 HTTP {exc.code}：{safe_error}",
                status=exc.code,
            ) from exc
        except (TimeoutError, socket.timeout) as exc:
            raise LLMRequestError("请求超时，已使用本地文字评分") from exc
        except error.URLError as exc:
            reason = exc.reason
            if isinstance(reason, (TimeoutError, socket.timeout)):
                raise LLMRequestError(
                    "请求超时，已使用本地文字评分"
                ) from exc
            safe_reason = redact_sensitive_text(
                reason,
                secrets=(self.api_key,),
            )
            raise LLMRequestError(
                f"无法连接大模型服务：{safe_reason}"
            ) from exc
        except OSError as exc:
            safe_error = redact_sensitive_text(
                exc,
                secrets=(self.api_key,),
            )
            raise LLMRequestError(
                f"大模型网络请求失败：{safe_error}"
            ) from exc
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise LLMRequestError("服务返回的外层响应不是有效 JSON") from exc
        if not isinstance(parsed, dict):
            raise LLMRequestError("服务返回的外层响应结构不正确")
        return parsed

    def list_models(self) -> list[str]:
        """读取服务端模型列表，失败时返回空列表并保留原因。"""
        try:
            response = self._request_json("GET", "/models")
            data = response.get("data", [])
            if not isinstance(data, list):
                raise LLMRequestError("模型列表结构不正确")
            models = [
                str(item["id"])
                for item in data
                if isinstance(item, dict) and item.get("id")
            ]
            self.last_error = ""
            return models
        except LLMRequestError as exc:
            self.last_error = redact_sensitive_text(
                exc,
                secrets=(self.api_key,),
            )
            return []

    def test_connection(self) -> tuple[bool, str]:
        """测试模型列表端点，但不发送任何业务数据。"""
        models = self.list_models()
        if self.last_error:
            return False, self.last_error
        return True, f"连接成功，发现 {len(models)} 个模型"

    def _prompt(
        self,
        payload: dict[str, Any],
        *,
        repair_text: str = "",
    ) -> str:
        instruction = (
            "你只能从给定 candidate_id 中选择一个最符合文字语义的候选。"
            "不得计算或修改金额、方向、日期、重要性阈值和最终处理状态。"
            "请严格返回指定 JSON 字段。"
        )
        if repair_text:
            instruction += (
                "上一次回答格式不合格，请只修复格式，不新增候选。"
                f"上次回答：{repair_text}"
            )
        return (
            instruction
            + "\n候选数据："
            + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        )

    def _protocol_payload(
        self,
        protocol: str,
        prompt: str,
    ) -> dict[str, Any]:
        schema = _decision_schema()
        if protocol == "responses":
            return {
                "model": self.config.model,
                "input": [
                    {
                        "role": "system",
                        "content": (
                            "你是银行流水核对的文字语义辅助工具，"
                            "只返回结构化判断。"
                        ),
                    },
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0,
                "text": {
                    "format": {
                        "type": "json_schema",
                        "name": "candidate_semantic_decision",
                        "strict": True,
                        "schema": schema,
                    }
                },
            }
        return {
            "model": self.config.model,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "你是银行流水核对的文字语义辅助工具，"
                        "只返回结构化判断。"
                    ),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
            "response_format": {
                "type": "json_schema",
                "json_schema": {
                    "name": "candidate_semantic_decision",
                    "strict": True,
                    "schema": schema,
                },
            },
        }

    def _call_protocol(
        self,
        protocol: str,
        prompt: str,
    ) -> dict[str, Any]:
        path = (
            "/responses"
            if protocol == "responses"
            else "/chat/completions"
        )
        return self._request_json(
            "POST",
            path,
            self._protocol_payload(protocol, prompt),
        )

    @staticmethod
    def _extract_model_text(
        protocol: str,
        response: dict[str, Any],
    ) -> str:
        if protocol == "responses":
            direct = response.get("output_text")
            if isinstance(direct, str):
                return direct
            for output in response.get("output", []):
                if not isinstance(output, dict):
                    continue
                for content in output.get("content", []):
                    if isinstance(content, dict) and isinstance(
                        content.get("text"),
                        str,
                    ):
                        return content["text"]
            raise ValueError("Responses 响应中缺少文字结果")
        choices = response.get("choices", [])
        if not choices or not isinstance(choices[0], dict):
            raise ValueError("Chat Completions 响应中缺少 choices")
        message = choices[0].get("message", {})
        content = message.get("content") if isinstance(message, dict) else None
        if not isinstance(content, str):
            raise ValueError("Chat Completions 响应中缺少文字结果")
        return content

    @staticmethod
    def _parse_decision(
        raw_text: str,
        candidate_ids: set[str],
    ) -> dict[str, Any]:
        cleaned = raw_text.strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
            cleaned = re.sub(r"\s*```$", "", cleaned)
        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError as exc:
            raise ValueError("回答不是有效 JSON") from exc
        if not isinstance(parsed, dict):
            raise ValueError("回答必须是 JSON 对象")
        missing = REQUIRED_DECISION_KEYS - parsed.keys()
        if missing:
            raise ValueError(f"回答缺少字段：{','.join(sorted(missing))}")
        selected = parsed["selected_candidate_id"]
        if not isinstance(selected, str) or selected not in candidate_ids:
            raise ValueError("大模型返回了候选集外编号")
        score = parsed["semantic_score"]
        if isinstance(score, bool) or not isinstance(score, int):
            raise ValueError("语义评分必须是整数")
        if not 0 <= score <= 100:
            raise ValueError("语义评分必须在零到一百之间")
        for key in ("supporting_evidence", "conflicting_evidence"):
            values = parsed[key]
            if not isinstance(values, list) or not all(
                isinstance(value, str) for value in values
            ):
                raise ValueError(f"{key} 必须是文字列表")
        for key in ("reason", "uncertainty", "suggested_status"):
            if not isinstance(parsed[key], str):
                raise ValueError(f"{key} 必须是文字")
        if parsed["suggested_status"] not in ALLOWED_SUGGESTED_STATUSES:
            raise ValueError("建议状态不在允许范围内")
        return parsed

    @staticmethod
    def _usage(response: dict[str, Any]) -> dict[str, int]:
        usage = response.get("usage", {})
        if not isinstance(usage, dict):
            return {}
        return {
            str(key): int(value)
            for key, value in usage.items()
            if isinstance(value, int) and not isinstance(value, bool)
        }

    def _fallback_record(
        self,
        semantic_request: CandidateSemanticRequest,
        *,
        started_at: str,
        start_time: float,
        error_message: object,
        raw_response: str = "",
        candidate_ids: tuple[str, ...] = (),
        sent_fields: tuple[str, ...] = (),
        protocol: str = "",
    ) -> LLMDecisionRecord:
        return LLMDecisionRecord(
            request_id=semantic_request.request_id,
            candidate_ids=candidate_ids,
            sent_fields=sent_fields,
            protocol=protocol,
            provider=self.provider_name,
            model=self.config.model,
            started_at=started_at,
            duration_ms=max(0, round((time.perf_counter() - start_time) * 1000)),
            fallback_used=True,
            error=redact_sensitive_text(
                error_message,
                secrets=(self.api_key,),
            ),
            raw_response=redact_sensitive_text(
                raw_response,
                secrets=(self.api_key,),
            ),
        )

    def evaluate_candidates(
        self,
        semantic_request: CandidateSemanticRequest,
    ) -> LLMDecisionRecord:
        """判断已有候选；任何异常都转成可追溯的本地降级记录。"""
        start_time = time.perf_counter()
        started_at = datetime.now().astimezone().isoformat(timespec="seconds")
        if not self.config.enabled:
            return self._fallback_record(
                semantic_request,
                started_at=started_at,
                start_time=start_time,
                error_message="大模型辅助未启用",
            )
        if not self.config.model:
            return self._fallback_record(
                semantic_request,
                started_at=started_at,
                start_time=start_time,
                error_message="尚未选择大模型",
            )

        limited_candidates = semantic_request.candidates[
            : max(0, self.config.candidate_limit)
        ]
        if not limited_candidates:
            return self._fallback_record(
                semantic_request,
                started_at=started_at,
                start_time=start_time,
                error_message="没有可供大模型判断的候选",
            )
        payload = build_online_payload(
            CandidateSemanticRequest(
                request_id=semantic_request.request_id,
                candidates=limited_candidates,
            ),
            mode=self.config.mode,
            local_fields=self.config.local_fields,
        )
        submitted_candidate_ids = tuple(
            item.candidate_id for item in limited_candidates
        )
        submitted_fields = {"日期", "金额"}
        for item in payload["candidates"]:
            submitted_fields.update(item["bank_fields"])
            submitted_fields.update(item["journal_fields"])
            if item["local_signals"]:
                submitted_fields.add("本机敏感字段一致性信号")
        submitted_fields_tuple = tuple(sorted(submitted_fields))
        candidate_ids = {
            candidate.candidate_id for candidate in limited_candidates
        }
        protocol_name = self.config.protocol.lower().strip()
        if protocol_name in {"chat", "chat_completion", "chat-completions"}:
            protocol_name = "chat_completions"
        protocols = (
            ["responses", "chat_completions"]
            if protocol_name == "auto"
            else [protocol_name]
        )
        if any(
            protocol not in {"responses", "chat_completions"}
            for protocol in protocols
        ):
            return self._fallback_record(
                semantic_request,
                started_at=started_at,
                start_time=start_time,
                error_message="调用协议配置不正确",
            )

        last_error: object = "大模型调用失败"
        last_raw = ""
        for protocol_index, protocol in enumerate(protocols):
            prompt = self._prompt(payload)
            try:
                response = self._call_protocol(protocol, prompt)
            except LLMRequestError as exc:
                last_error = exc
                can_fallback_protocol = (
                    protocol_name == "auto"
                    and protocol_index == 0
                    and exc.status in {400, 404, 405, 501}
                )
                if can_fallback_protocol:
                    continue
                return self._fallback_record(
                    semantic_request,
                    started_at=started_at,
                    start_time=start_time,
                    error_message=exc,
                    candidate_ids=submitted_candidate_ids,
                    sent_fields=submitted_fields_tuple,
                    protocol=protocol,
                )

            try:
                raw_text = self._extract_model_text(protocol, response)
                last_raw = raw_text
                parsed = self._parse_decision(raw_text, candidate_ids)
            except ValueError as first_error:
                repair_prompt = self._prompt(
                    payload,
                    repair_text=redact_sensitive_text(
                        last_raw or first_error,
                        secrets=(self.api_key,),
                    ),
                )
                try:
                    response = self._call_protocol(protocol, repair_prompt)
                    raw_text = self._extract_model_text(protocol, response)
                    last_raw = raw_text
                    parsed = self._parse_decision(raw_text, candidate_ids)
                except (LLMRequestError, ValueError) as second_error:
                    return self._fallback_record(
                        semantic_request,
                        started_at=started_at,
                        start_time=start_time,
                        error_message=second_error,
                        raw_response=last_raw,
                        candidate_ids=submitted_candidate_ids,
                        sent_fields=submitted_fields_tuple,
                        protocol=protocol,
                    )

            return LLMDecisionRecord(
                request_id=semantic_request.request_id,
                candidate_ids=submitted_candidate_ids,
                sent_fields=submitted_fields_tuple,
                protocol=protocol,
                selected_candidate_id=parsed["selected_candidate_id"],
                semantic_score=parsed["semantic_score"],
                reason=parsed["reason"],
                supporting_evidence=tuple(parsed["supporting_evidence"]),
                conflicting_evidence=tuple(parsed["conflicting_evidence"]),
                uncertainty=parsed["uncertainty"],
                suggested_status=parsed["suggested_status"],
                provider=self.provider_name,
                model=self.config.model,
                started_at=started_at,
                duration_ms=max(
                    0,
                    round((time.perf_counter() - start_time) * 1000),
                ),
                usage=self._usage(response),
                fallback_used=False,
                raw_response=redact_sensitive_text(
                    last_raw,
                    secrets=(self.api_key,),
                ),
            )

        return self._fallback_record(
            semantic_request,
            started_at=started_at,
            start_time=start_time,
            error_message=last_error,
            raw_response=last_raw,
            candidate_ids=submitted_candidate_ids,
            sent_fields=submitted_fields_tuple,
            protocol=protocols[-1] if protocols else "",
        )
