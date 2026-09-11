"""Application configuration.

Configuration can be provided via environment variables (or a ``.env`` file
in the working directory) and can be overridden from the command line. See
``.env.example`` for the full list of supported variables.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

try:
    from dotenv import load_dotenv

    # Load variables from a .env file if present, without overriding ones
    # already set in the real environment.
    load_dotenv(override=False)
except ImportError:  # pragma: no cover - python-dotenv is a declared dependency
    pass


def _get_float(name: str, default: float) -> float:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    try:
        return float(value)
    except ValueError:
        return default


def _get_int(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    try:
        return int(value)
    except ValueError:
        return default


def _get_str(name: str, default: str) -> str:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value


def _get_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None or value == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


@dataclass
class Settings:
    """Global application configuration (with sensible defaults)."""

    provider: str = _get_str("TRANSLATOR_PROVIDER", "local")
    remove_audio: bool = _get_bool("TRANSLATOR_REMOVE_AUDIO", False)
    max_chars_per_request: int = _get_int("TRANSLATOR_MAX_CHARS_PER_REQUEST", 4500)
    request_delay_seconds: float = _get_float("TRANSLATOR_REQUEST_DELAY", 0.4)
    max_retries: int = _get_int("TRANSLATOR_MAX_RETRIES", 4)
    retry_backoff_seconds: float = _get_float("TRANSLATOR_RETRY_BACKOFF", 2.0)
    request_timeout_seconds: float = _get_float("TRANSLATOR_TIMEOUT", 60.0)
    api_key: str | None = (
        os.getenv("TRANSLATOR_API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or None
    )
    api_base_url: str | None = (
        os.getenv("TRANSLATOR_API_BASE_URL")
        or os.getenv("OPENAI_API_BASE_URL")
        or None
    )
    model: str | None = (
        os.getenv("TRANSLATOR_MODEL")
        or os.getenv("OPENAI_MODEL")
        or None
    )
    log_level: str | None = (os.getenv("TRANSLATOR_LOG_LEVEL") or None)
    cache_dir: Path | None = (
        Path(os.getenv("PPTX_TRANSLATOR_CACHE_DIR"))
        if os.getenv("PPTX_TRANSLATOR_CACHE_DIR")
        else None
    )


def load_settings() -> Settings:
    """Returns the configuration loaded from the environment."""

    return Settings()
