"""配置加载：从 ``.env`` 读取模型接入信息。

约定（全部可选，均可用环境变量覆盖）::

    PPT_KING_API_KEY=sk-xxx
    PPT_KING_BASE_URL=https://api.deepseek.com
    PPT_KING_MODEL=deepseek-flash
    PPT_KING_TEMPERATURE=0
    PPT_KING_MAX_TOKENS=8000
    PPT_KING_TIMEOUT=120
    PPT_KING_MAX_RETRIES=3
    PPT_KING_CONCURRENCY=4

兼容写法：若 ``.env`` 中直接写一行 ``sk-...``（没有键名），也视作 API Key。
"""

from __future__ import annotations

import os
from dataclasses import dataclass, replace
from pathlib import Path

DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-flash"

_ALIASES = {
    "api_key": ("PPT_KING_API_KEY", "DEEPSEEK_API_KEY", "OPENAI_API_KEY"),
    "base_url": ("PPT_KING_BASE_URL", "DEEPSEEK_BASE_URL", "OPENAI_BASE_URL"),
    "model": ("PPT_KING_MODEL", "DEEPSEEK_MODEL", "OPENAI_MODEL"),
}


class ConfigurationError(RuntimeError):
    """配置缺失或不合法。"""


@dataclass(frozen=True)
class Settings:
    api_key: str
    base_url: str = DEFAULT_BASE_URL
    model: str = DEFAULT_MODEL
    temperature: float = 0.0
    max_tokens: int = 8000
    timeout: float = 120.0
    max_retries: int = 3
    concurrency: int = 4
    cache_dir: Path | None = None
    env_file: Path | None = None

    def describe(self) -> str:
        return f"model={self.model} base_url={self.base_url}"


def _parse_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    bare: list[str] = []
    for raw in path.read_text(encoding="utf-8-sig", errors="replace").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
        else:
            bare.append(line)
    if "PPT_KING_API_KEY" not in values:
        for item in bare:
            if item.startswith("sk-"):
                values["PPT_KING_API_KEY"] = item
                break
    return values


def _find_env_file(start: Path | None = None) -> Path | None:
    current = (start or Path.cwd()).resolve()
    for candidate in (current, *current.parents):
        env_path = candidate / ".env"
        if env_path.is_file():
            return env_path
    return None


def _first(values: dict[str, str], keys: tuple[str, ...]) -> str | None:
    for key in keys:
        value = values.get(key) or os.environ.get(key)
        if value:
            return value.strip()
    return None


def _as_int(values: dict[str, str], key: str, default: int) -> int:
    raw = values.get(key) or os.environ.get(key)
    if not raw:
        return default
    try:
        return int(float(raw))
    except ValueError as exc:
        raise ConfigurationError(f"{key} 不是合法整数: {raw!r}") from exc


def _as_float(values: dict[str, str], key: str, default: float) -> float:
    raw = values.get(key) or os.environ.get(key)
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError as exc:
        raise ConfigurationError(f"{key} 不是合法数字: {raw!r}") from exc


def load_settings(
    env_file: Path | str | None = None,
    *,
    overrides: dict[str, object] | None = None,
    require_api_key: bool = True,
) -> Settings:
    """读取配置。``overrides`` 优先级最高。"""
    path = Path(env_file) if env_file else _find_env_file()
    values = _parse_env_file(path) if path and path.is_file() else {}

    api_key = _first(values, _ALIASES["api_key"]) or ""
    settings = Settings(
        api_key=api_key,
        base_url=_first(values, _ALIASES["base_url"]) or DEFAULT_BASE_URL,
        model=_first(values, _ALIASES["model"]) or DEFAULT_MODEL,
        temperature=_as_float(values, "PPT_KING_TEMPERATURE", 0.0),
        max_tokens=_as_int(values, "PPT_KING_MAX_TOKENS", 8000),
        timeout=_as_float(values, "PPT_KING_TIMEOUT", 120.0),
        max_retries=_as_int(values, "PPT_KING_MAX_RETRIES", 3),
        concurrency=_as_int(values, "PPT_KING_CONCURRENCY", 4),
        cache_dir=Path(values["PPT_KING_CACHE_DIR"])
        if values.get("PPT_KING_CACHE_DIR")
        else None,
        env_file=path,
    )

    if overrides:
        clean = {key: value for key, value in overrides.items() if value not in (None, "")}
        settings = replace(settings, **clean)

    if require_api_key and not settings.api_key:
        hint = (
            f"请在 {path} 中填写 PPT_KING_API_KEY" if path else "请在项目根目录创建 .env 并填写 PPT_KING_API_KEY"
        )
        raise ConfigurationError(f"未配置模型 API Key。{hint}")
    if settings.concurrency < 1:
        raise ConfigurationError("PPT_KING_CONCURRENCY 必须 >= 1")
    return settings
