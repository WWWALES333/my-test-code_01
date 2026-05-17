from __future__ import annotations

import json
import os
import re
import subprocess
import time
from dataclasses import dataclass
from typing import Dict, Iterable, List

import requests


DEFAULT_BASE_URL = "https://api.minimaxi.com/v1"
DEFAULT_MODEL = "MiniMax-M2.7"
KEYCHAIN_SERVICE = "my-test-code_01/minimax"
KEYCHAIN_ACCOUNT = "minimax_api_key"


@dataclass
class ModelConfig:
    api_key: str
    base_url: str = DEFAULT_BASE_URL
    model: str = DEFAULT_MODEL
    timeout_seconds: int = 60
    max_retries: int = 2


class OpenAICompatibleClient:
    """OpenAI 兼容模型客户端，默认用于 Minimax。

    密钥只从环境变量或 macOS Keychain 读取，不写入仓库和运行产物。
    """

    def __init__(self, config: ModelConfig | None = None) -> None:
        self.config = config or load_model_config()

    def chat_json(
        self,
        messages: List[Dict[str, str]],
        *,
        temperature: float = 0,
        max_tokens: int = 800,
    ) -> Dict[str, object]:
        content = self.chat_text(messages, temperature=temperature, max_tokens=max_tokens)
        return parse_json_payload(content)

    def chat_text(
        self,
        messages: List[Dict[str, str]],
        *,
        temperature: float = 0,
        max_tokens: int = 800,
    ) -> str:
        if not self.config.api_key:
            raise RuntimeError("OPENAI_API_KEY 未配置，且 Keychain 未找到 Minimax key")

        url = f"{self.config.base_url.rstrip('/')}/chat/completions"
        payload = {
            "model": self.config.model,
            "temperature": temperature,
            "max_tokens": max_tokens,
            "messages": messages,
        }
        headers = {
            "Authorization": f"Bearer {self.config.api_key}",
            "Content-Type": "application/json",
        }
        last_error: Exception | None = None
        for attempt in range(self.config.max_retries + 1):
            try:
                response = requests.post(url, headers=headers, json=payload, timeout=self.config.timeout_seconds)
                response.raise_for_status()
                data = response.json()
                content = (
                    data.get("choices", [{}])[0]
                    .get("message", {})
                    .get("content", "")
                )
                if not content:
                    raise RuntimeError("模型返回为空")
                return str(content)
            except Exception as exc:  # noqa: BLE001 - 需要把 HTTP/JSON/超时统一兜底
                last_error = exc
                if attempt < self.config.max_retries:
                    time.sleep(0.8 * (attempt + 1))
                    continue
                raise RuntimeError(f"模型调用失败: {type(exc).__name__}: {str(exc)[:180]}") from exc
        raise RuntimeError(f"模型调用失败: {last_error}")


def load_model_config() -> ModelConfig:
    api_key = os.getenv("OPENAI_API_KEY", "").strip() or _read_keychain_secret()
    return ModelConfig(
        api_key=api_key,
        base_url=os.getenv("OPENAI_BASE_URL", DEFAULT_BASE_URL).strip().rstrip("/") or DEFAULT_BASE_URL,
        model=os.getenv("OPENAI_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL,
        timeout_seconds=int(os.getenv("OPENAI_TIMEOUT_SECONDS", "60")),
        max_retries=int(os.getenv("OPENAI_MAX_RETRIES", "2")),
    )


def parse_json_payload(content: str) -> Dict[str, object]:
    """清洗 Minimax `<think>` 输出并提取首个可解析 JSON 对象。"""
    cleaned = strip_think_blocks(content).strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        parsed = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        extracted = extract_json_object(cleaned)
        if extracted is None:
            raise RuntimeError(f"模型返回非 JSON: {exc}") from exc
        try:
            parsed = json.loads(extracted)
        except json.JSONDecodeError as inner_exc:
            raise RuntimeError(f"模型返回非 JSON: {inner_exc}") from inner_exc
    if not isinstance(parsed, dict):
        raise RuntimeError("模型返回结构不是 JSON 对象")
    return parsed


def strip_think_blocks(content: str) -> str:
    return re.sub(r"<think>.*?</think>", "", content, flags=re.S | re.I).strip()


def extract_json_object(content: str) -> str | None:
    start = content.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escaped = False
    for idx in range(start, len(content)):
        ch = content[idx]
        if escaped:
            escaped = False
            continue
        if ch == "\\":
            escaped = True
            continue
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return content[start : idx + 1]
    return None


def _read_keychain_secret() -> str:
    try:
        value = subprocess.check_output(
            [
                "security",
                "find-generic-password",
                "-s",
                KEYCHAIN_SERVICE,
                "-a",
                KEYCHAIN_ACCOUNT,
                "-w",
            ],
            text=True,
            stderr=subprocess.DEVNULL,
        )
        return value.strip()
    except Exception:
        return ""
