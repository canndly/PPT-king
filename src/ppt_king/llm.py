"""OpenAI 兼容接口客户端（DeepSeek / OpenAI / vLLM 等）。"""

from __future__ import annotations

import threading

from openai import (
    APIConnectionError,
    APIStatusError,
    APITimeoutError,
    OpenAI,
    RateLimitError,
)
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

from .config import Settings


class LLMError(RuntimeError):
    """模型调用失败，且不可恢复。"""


class RetryableError(Exception):
    """需要退避重试的临时错误。"""


def _is_retryable(exc: BaseException) -> bool:
    if isinstance(exc, (RateLimitError, APIConnectionError, APITimeoutError, RetryableError)):
        return True
    if isinstance(exc, APIStatusError):
        return exc.status_code >= 500 or exc.status_code == 429
    return False


class LLMClient:
    def __init__(self, settings: Settings):
        self.settings = settings
        self._lock = threading.Lock()
        self._client = OpenAI(
            api_key=settings.api_key,
            base_url=settings.base_url,
            timeout=settings.timeout,
            max_retries=0,  # 重试由 tenacity 统一控制
        )
        self.calls = 0
        self.prompt_tokens = 0
        self.completion_tokens = 0

    # -- 内部 ---------------------------------------------------------------
    def _request(self, system: str, user: str) -> str:
        @retry(
            retry=retry_if_exception_type(RetryableError),
            stop=stop_after_attempt(max(1, self.settings.max_retries)),
            wait=wait_exponential(multiplier=1, min=2, max=30),
            reraise=True,
        )
        def _call() -> str:
            try:
                response = self._client.chat.completions.create(
                    model=self.settings.model,
                    messages=[
                        {"role": "system", "content": system},
                        {"role": "user", "content": user},
                    ],
                    temperature=self.settings.temperature,
                    max_tokens=self.settings.max_tokens,
                )
            except LLMError:
                raise
            except Exception as exc:  # noqa: BLE001 - 按状态码决定是否重试
                if _is_retryable(exc):
                    raise RetryableError(str(exc)) from exc
                raise LLMError(f"模型调用失败：{exc}") from exc

            if response.usage:
                with self._lock:
                    self.prompt_tokens += response.usage.prompt_tokens or 0
                    self.completion_tokens += response.usage.completion_tokens or 0
            if not response.choices:
                raise LLMError("模型返回空 choices")
            content = response.choices[0].message.content
            if not content or not content.strip():
                raise RetryableError("模型返回空内容")
            return content

        try:
            return _call()
        except LLMError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise LLMError(f"调用重试 {self.settings.max_retries} 次后仍失败：{exc}") from exc

    # -- 对外 ---------------------------------------------------------------
    def complete(self, system: str, user: str) -> str:
        with self._lock:
            self.calls += 1
        return self._request(system, user)

    def stats(self) -> str:
        return (
            f"调用 {self.calls} 次，输入 {self.prompt_tokens} tokens，"
            f"输出 {self.completion_tokens} tokens"
        )
