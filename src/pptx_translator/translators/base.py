"""Base interface for translation providers."""

from __future__ import annotations

import abc
import logging
import time
from typing import Iterable

logger = logging.getLogger(__name__)


class TranslationError(RuntimeError):
    """Unrecoverable error while translating text with a provider."""


class BaseTranslator(abc.ABC):
    """Base class providing retries, rate limiting and caching.

    Subclasses only need to implement :meth:`_translate_one`. This class
    takes care of:

    * Not repeating calls for texts already translated (in-memory cache).
    * Spacing out API calls (``request_delay_seconds``) to avoid abusing
      the service and getting the application blocked.
    * Retrying with exponential backoff on transient errors.
    """

    name = "base"

    def __init__(
        self,
        request_delay_seconds: float = 0.4,
        max_retries: int = 4,
        retry_backoff_seconds: float = 2.0,
    ) -> None:
        self.request_delay_seconds = request_delay_seconds
        self.max_retries = max_retries
        self.retry_backoff_seconds = retry_backoff_seconds
        self._cache: dict[tuple[str, str, str, str | None], str] = {}
        self._last_request_ts: float = 0.0

    @abc.abstractmethod
    def _translate_one(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        context: str | None = None,
    ) -> str:
        """Translates a single string. Must be implemented by each provider."""

    def on_presentation_start(self, context: str | None = None) -> None:
        """Hook called once when translating a presentation."""

    def _throttle(self) -> None:
        if self.request_delay_seconds <= 0:
            return
        elapsed = time.monotonic() - self._last_request_ts
        remaining = self.request_delay_seconds - elapsed
        if remaining > 0:
            time.sleep(remaining)

    def translate(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        context: str | None = None,
    ) -> str:
        """Translates ``text``, reusing the cache and retrying on failure."""

        if not text or not text.strip():
            return text

        cache_key = (text, source_lang, target_lang, context)
        if cache_key in self._cache:
            return self._cache[cache_key]

        last_error: Exception | None = None
        for attempt in range(1, self.max_retries + 1):
            try:
                self._throttle()
                try:
                    translated = self._translate_one(
                        text,
                        source_lang,
                        target_lang,
                        context=context,
                    )
                except TypeError as exc:
                    if "context" not in str(exc):
                        raise
                    logger.debug(
                        "[%s] Provider does not accept the optional context argument; retrying without context.",
                        self.name,
                    )
                    translated = self._translate_one(text, source_lang, target_lang)
                self._last_request_ts = time.monotonic()
                if translated is None or translated == "":
                    translated = text
                self._cache[cache_key] = translated
                return translated
            except Exception as exc:  # noqa: BLE001 - we want to retry any transient failure
                last_error = exc
                wait = self.retry_backoff_seconds * attempt
                logger.warning(
                    "[%s] Translation failed (attempt %d/%d): %s. Retrying in %.1fs...",
                    self.name,
                    attempt,
                    self.max_retries,
                    exc,
                    wait,
                )
                time.sleep(wait)

        raise TranslationError(
            f"Could not translate the text after {self.max_retries} attempts: {last_error}"
        )

    def translate_many(
        self,
        texts: Iterable[str],
        source_lang: str,
        target_lang: str,
        fallback: "BaseTranslator | None" = None,
        context: str | None = None,
    ) -> tuple[dict[str, str], list[str]]:
        """Translates a collection of unique texts.

        Returns a tuple ``(results, failed)`` where ``results`` is a
        dictionary ``{original_text: translated_text}`` and ``failed`` is
        the list of texts that could not be translated (not even with
        ``fallback``, if one is provided), and therefore keep the original
        text.
        """

        results: dict[str, str] = {}
        failed: list[str] = []
        texts = list(dict.fromkeys(texts))  # remove duplicates while preserving order
        total = len(texts)
        for index, text in enumerate(texts, start=1):
            try:
                results[text] = self.translate(text, source_lang, target_lang, context=context)
            except TranslationError as primary_exc:
                translated = None
                if fallback is not None:
                    logger.warning(
                        "[%s] Could not translate '%s...'; trying the fallback "
                        "provider '%s'.",
                        self.name,
                        text[:60],
                        fallback.name,
                    )
                    try:
                        translated = fallback.translate(text, source_lang, target_lang, context=context)
                    except TranslationError as fallback_exc:
                        logger.error(
                            "[%s] The fallback provider could not translate "
                            "'%s...' either: %s",
                            fallback.name,
                            text[:60],
                            fallback_exc,
                        )
                if translated is not None:
                    results[text] = translated
                else:
                    # Don't abort the whole presentation because a single text
                    # failed: keep the original text and continue with the rest.
                    logger.error(
                        "Could not translate '%s...'; keeping the original text. Cause: %s",
                        text[:60],
                        primary_exc,
                    )
                    results[text] = text
                    failed.append(text)
            if total >= 10 and index % 10 == 0:
                logger.info("Translated %d/%d unique texts...", index, total)
        return results, failed
