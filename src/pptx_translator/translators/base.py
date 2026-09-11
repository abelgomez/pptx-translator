"""Base interface for translation providers."""

from __future__ import annotations

import abc
import logging
import time
from typing import Any, Iterable, Sequence

logger = logging.getLogger(__name__)


class TranslationError(RuntimeError):
    """Unrecoverable error while translating text with a provider."""


class BaseTranslator(abc.ABC):
    """Common contract for all translation backends."""

    name = "base"

    def __init__(
        self,
        request_delay_seconds: float = 0.4,
        max_retries: int = 4,
        retry_backoff_seconds: float = 2.0,
        exception_rules: Sequence[Any] | None = None,
    ) -> None:
        self.request_delay_seconds = request_delay_seconds
        self.max_retries = max_retries
        self.retry_backoff_seconds = retry_backoff_seconds
        self.exception_rules = list(exception_rules) if exception_rules is not None else []
        self._cache: dict[tuple[str, str, str, str | None], str] = {}
        self._last_request_ts = 0.0

    @abc.abstractmethod
    def _translate_slide(
        self,
        texts: Iterable[str],
        source_lang: str,
        target_lang: str,
        context: str | None = None,
        exception_rules: Sequence[Any] | None = None,
    ) -> dict[str, str]:
        """Translate a batch of texts for one slide."""

    def _translate_single_text(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        context: str | None = None,
        exception_rules: Sequence[Any] | None = None,
    ) -> str:
        rules = exception_rules if exception_rules is not None else self.exception_rules
        translated_by_text = self._translate_slide(
            [text],
            source_lang,
            target_lang,
            context=context,
            exception_rules=rules,
        )
        return translated_by_text.get(text, text)

    def on_presentation_start(self, context: str | None = None) -> None:
        """Hook called once when translating a presentation."""

    def _throttle(self) -> None:
        if self.request_delay_seconds <= 0:
            return
        elapsed = time.monotonic() - self._last_request_ts
        remaining = self.request_delay_seconds - elapsed
        if remaining > 0:
            time.sleep(remaining)

    def translate_slide(
        self,
        texts: Iterable[str],
        source_lang: str,
        target_lang: str,
        fallback: "BaseTranslator | None" = None,
        context: str | None = None,
        exception_rules: Sequence[Any] | None = None,
    ) -> tuple[dict[str, str], list[str]]:
        """Translate a batch of texts. Return (results, failed)."""

        unique_texts = list(dict.fromkeys(texts))
        results: dict[str, str] = {}
        failed: list[str] = []
        rules = exception_rules if exception_rules is not None else self.exception_rules

        try:
            translated_by_text = self._translate_slide(
                unique_texts,
                source_lang,
                target_lang,
                context=context,
                exception_rules=rules,
            )
        except TranslationError:
            for text in unique_texts:
                try:
                    translated = self._translate_single_text(
                        text,
                        source_lang,
                        target_lang,
                        context=context,
                        exception_rules=rules,
                    )
                except TranslationError as inner_exc:
                    if self.name == "local":
                        logger.warning(
                            "[%s] Could not translate '%s...'; keeping the original text without fallback.",
                            self.name,
                            text[:60],
                        )
                        translated = text
                        failed.append(text)
                    elif fallback is not None:
                        logger.warning(
                            "[%s] Could not translate '%s...'; trying the fallback provider '%s'.",
                            self.name,
                            text[:60],
                            fallback.name,
                        )
                        try:
                            translated = fallback._translate_single_text(
                                text,
                                source_lang,
                                target_lang,
                                context=context,
                                exception_rules=fallback.exception_rules,
                            )
                        except TranslationError as fallback_exc:
                            logger.error(
                                "[%s] The fallback provider could not translate '%s...' either: %s",
                                fallback.name,
                                text[:60],
                                fallback_exc,
                            )
                            translated = text
                            failed.append(text)
                    else:
                        logger.error(
                            "Could not translate '%s...'; keeping the original text. Cause: %s",
                            text[:60],
                            inner_exc,
                        )
                        translated = text
                        failed.append(text)

                results[text] = translated
            return results, failed

        for text in unique_texts:
            results[text] = translated_by_text.get(text, text)
        return results, failed
