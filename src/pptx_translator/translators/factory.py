"""Translator factory: creates the appropriate instance based on configuration."""

from __future__ import annotations

from ..config import Settings
from .base import BaseTranslator
from .openai import OpenAITranslator
from .local_argos import LocalArgosTranslator

PROVIDERS = ("local", "openai")


def create_translator(
    settings: Settings,
    provider: str | None = None,
    api_key: str | None = None,
    api_base_url: str | None = None,
    model: str | None = None,
    request_delay_seconds: float | None = None,
    request_timeout_seconds: float | None = None,
) -> BaseTranslator:
    """Creates the configured translator."""

    provider_name = (provider or settings.provider or "local").strip().lower()

    kwargs = {
        "request_delay_seconds": (
            request_delay_seconds
            if request_delay_seconds is not None
            else settings.request_delay_seconds
        ),
        "max_retries": settings.max_retries,
        "retry_backoff_seconds": settings.retry_backoff_seconds,
    }

    if provider_name == "local":
        return LocalArgosTranslator(**kwargs)
    if provider_name == "openai":
        return OpenAITranslator(
            api_key=api_key or settings.api_key,
            api_base_url=api_base_url or settings.api_base_url,
            model=model or settings.model,
            request_timeout_seconds=(
                request_timeout_seconds
                if request_timeout_seconds is not None
                else settings.request_timeout_seconds
            ),
            **kwargs,
        )

    raise ValueError(
        f"Unknown translation provider: '{provider_name}'. "
        f"Valid options: {', '.join(PROVIDERS)}."
    )
