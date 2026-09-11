"""HTTP-based translation providers compatible with OpenAI-style APIs."""

from __future__ import annotations

import json
import logging
import os
import re
from typing import Any

import requests

from ..exceptions_list import ExceptionMode, ExceptionRule
from .base import BaseTranslator, TranslationError

logger = logging.getLogger(__name__)

_SUPPORTED_LANGUAGE_NAMES = {
    "es": "Spanish",
    "en": "English",
    "fr": "French",
    "de": "German",
    "it": "Italian",
    "pt": "Portuguese",
    "ca": "Catalan",
    "gl": "Galician",
    "eu": "Basque",
    "zh": "Chinese",
    "ja": "Japanese",
    "ko": "Korean",
    "ar": "Arabic",
    "ru": "Russian",
    "pl": "Polish",
    "nl": "Dutch",
    "sv": "Swedish",
}


def _language_name(code: str | None) -> str:
    if not code:
        return "target language"
    return _SUPPORTED_LANGUAGE_NAMES.get(code.lower(), code.upper())


def _format_context_for_log(context: str | None) -> str:
    if not context:
        return ""

    normalized = context.strip().replace("\r\n", "\n").replace("\r", "\n")
    normalized = re.sub(r"\n+", "\n", normalized)
    lines = [line.strip() for line in normalized.split("\n")]
    return " / ".join(line for line in lines if line)


class OpenAITranslator(BaseTranslator):
    """Translator backed by an OpenAI-compatible HTTP endpoint."""

    name = "openai"

    def __init__(
        self,
        api_key: str | None = None,
        api_base_url: str | None = None,
        model: str | None = None,
        request_timeout_seconds: float | None = None,
        *args,
        **kwargs,
    ) -> None:
        self.api_key = (api_key or os.getenv("TRANSLATOR_API_KEY") or os.getenv("OPENAI_API_KEY") or "").strip()
        self.api_base_url = (api_base_url or os.getenv("TRANSLATOR_API_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
        self.model = model or os.getenv("TRANSLATOR_MODEL") or "gpt-4o-mini"
        timeout_value = request_timeout_seconds
        if timeout_value is None:
            timeout_value = os.getenv("TRANSLATOR_TIMEOUT")
        try:
            self.request_timeout_seconds = float(timeout_value) if timeout_value not in (None, "") else 60.0
        except (TypeError, ValueError):
            self.request_timeout_seconds = 60.0
        self._protected_replacements: dict[str, str] = {}
        if not self.api_key:
            raise ValueError(
                "An API key is required for remote translation providers. "
                "Set TRANSLATOR_API_KEY or pass api_key=... ."
            )
        super().__init__(*args, **kwargs)

    def on_presentation_start(self, context: str | None = None) -> None:
        logger.info(
            "Using first-slide presentation context for OpenAI-compatible translation: %s",
            _format_context_for_log(context),
        )

    def _extract_text(self, payload: dict[str, Any]) -> str:
        choices = payload.get("choices") or []
        if not choices:
            raise TranslationError(f"Empty translation response: {payload}")

        first = choices[0]
        message = first.get("message") or {}
        content = message.get("content")
        if isinstance(content, list):
            parts: list[str] = []
            for item in content:
                if isinstance(item, dict):
                    text = item.get("text")
                    if isinstance(text, str):
                        parts.append(text)
                elif isinstance(item, str):
                    parts.append(item)
            content = "".join(parts)

        if not isinstance(content, str):
            raise TranslationError(f"Unexpected response format: {payload}")

        return content.strip()

    @staticmethod
    def _describe_exception_rules(exception_rules: list[ExceptionRule] | None) -> str:
        if not exception_rules:
            return ""

        entries: list[str] = []
        seen: set[str] = set()
        for rule in exception_rules:
            desc = f"{rule.raw_source} -> {rule.raw_destination}"
            if desc in seen:
                continue
            seen.add(desc)
            entries.append(desc)

        if not entries:
            return ""

        return (
            "The following exception terms must be kept exactly as specified in the target language: "
            + "; ".join(entries)
            + ". "
        )

    def _build_system_content(
        self,
        source_lang: str = "auto",
        target_lang: str | None = None,
        context: str | None = None,
        protected_replacements: dict[str, str] | None = None,
        exception_rules: list[ExceptionRule] | None = None,
    ) -> str:
        if target_lang is None and context is None and source_lang and source_lang not in {"auto", "es", "en", "fr", "de", "it", "pt", "ca", "gl", "eu", "zh", "ja", "ko", "ar", "ru", "pl", "nl", "sv"}:
            context = source_lang
            source_lang = "auto"
            target_lang = "auto"
        if target_lang is None:
            target_lang = "auto"

        effective_rules = exception_rules if exception_rules is not None else self.exception_rules
        protected_replacements = protected_replacements or self._protected_replacements
        replacement_text = self._describe_exception_rules(effective_rules)
        if not replacement_text and protected_replacements:
            seen_values: set[str] = set()
            protected_entries: list[str] = []
            for token, value in sorted(
                protected_replacements.items(),
                key=lambda item: (len(item[0]), item[1]),
                reverse=True,
            ):
                if value in seen_values:
                    continue
                seen_values.add(value)
                protected_entries.append(f"{token} -> {value}")
            replacement_text = (
                "Protected identifiers that must remain exactly unchanged and must be restored verbatim are: "
                + "; ".join(protected_entries)
                + ". "
            )

        source_name = _language_name(source_lang)
        target_name = _language_name(target_lang)
        exception_guidance = (
            "When the source text contains any of the exception terms listed above, keep that term exactly as the corresponding target value specifies. "
            "Do not translate, rearrange, or normalize it. Preserve acronyms, technical labels, and fixed names verbatim. "
        )
        base = (
            "You are a specialist technical translator for PowerPoint slides and presentation materials. "
            f"Translate from {source_name} to {target_name}. "
            "Translate every segment completely, while respecting the exception list exactly. "
            "Do not leave ordinary words untranslated in the source language. Preserve punctuation, casing, numbers, spacing, and the meaning of technical terms. "
            "Return the result as a fenced JSON code block with a single key 'translations'. "
            "Each entry must have exactly two fields: 'id' and 'text'. The list order must match the input order. "
            + replacement_text
            + exception_guidance
        )
        if not context:
            return base

        safe_context = " ".join(context.replace("\r", " ").replace("\n", " ").split())
        return (
            f"Presentation context: {safe_context}. "
            "Use this context to preserve the correct domain terminology, title conventions, author names, affiliations, and presentation-specific phrasing. "
            + base
        )

    def _restore_protected_replacements(self, text: str) -> str:
        replacements = getattr(self, "_protected_replacements", {})
        if not replacements:
            return text

        restored = text
        for token, value in sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True):
            restored = re.sub(re.escape(token), lambda _m, v=value: v, restored, flags=re.IGNORECASE)
        return restored

    @staticmethod
    def _strip_json_code_fence(raw: str) -> str:
        text = raw.strip()
        if not text.startswith("```"):
            return text

        match = re.search(r"```(?:json)?\s*(.*?)\s*```", text, flags=re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1).strip()
        return text.removeprefix("```").removesuffix("```").strip()

    def _request_translation(
        self,
        user_content: str,
        source_lang: str,
        target_lang: str,
        context: str | None = None,
        exception_rules: list[ExceptionRule] | None = None,
    ) -> str:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": self._build_system_content(
                        source_lang,
                        target_lang,
                        context,
                        getattr(self, "_protected_replacements", {}),
                        exception_rules,
                    ),
                },
                {"role": "user", "content": user_content},
            ],
            "temperature": 0,
        }

        logger.info(
            "Invoking OpenAI-compatible API (model=%s, source=%s, target=%s)",
            self.model,
            source_lang,
            target_lang,
        )
        logger.debug("OpenAI request payload:\n%s", json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True))

        response = requests.post(
            f"{self.api_base_url}/chat/completions",
            headers=headers,
            json=payload,
            timeout=self.request_timeout_seconds,
        )
        if response.status_code >= 400:
            raise TranslationError(
                f"HTTP {response.status_code} from remote provider: {response.text}"
            )

        try:
            data = response.json()
        except ValueError as exc:
            raise TranslationError(f"Invalid JSON response from remote provider: {response.text}") from exc

        return self._extract_text(data)

    def _translate_single_text(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        context: str | None = None,
        exception_rules: list[ExceptionRule] | None = None,
    ) -> str:
        rules = exception_rules if exception_rules is not None else self.exception_rules
        user_content = (
            f"Translate the following text from {_language_name(source_lang)} to {_language_name(target_lang)}. "
            "Translate the entire text fully and naturally. Do not leave ordinary words in the source language untranslated. "
            "Preserve punctuation, casing, numbers, and spacing. Keep acronyms, names, and protected identifiers unchanged only when they are explicitly required. "
            "Return only the translated text and nothing else.\n\n"
            f"{text}"
        )
        translated_text = self._request_translation(
            user_content,
            source_lang,
            target_lang,
            context=context,
            exception_rules=rules,
        )
        cleaned = translated_text.strip()
        if not cleaned:
            raise TranslationError("OpenAI returned an empty translation for a single text")
        return self._restore_protected_replacements(cleaned)

    def _translate_slide(
        self,
        texts: list[str],
        source_lang: str,
        target_lang: str,
        context: str | None = None,
        exception_rules: list[ExceptionRule] | None = None,
    ) -> dict[str, str]:
        """Translate every text fragment from a slide in a single API call."""

        rules = exception_rules if exception_rules is not None else self.exception_rules
        unique_texts = list(dict.fromkeys(texts))
        if not unique_texts:
            return {}

        entries_payload = [{"id": text, "text": text} for text in unique_texts]
        user_content = (
            f"Source language: {_language_name(source_lang)}\n"
            f"Target language: {_language_name(target_lang)}\n\n"
            "Translate every item below. Do not leave ordinary words untranslated. "
            "Only preserve values unchanged if they are explicitly protected exceptions or proper nouns/acronyms that must stay as-is. "
            "Return a fenced JSON code block with a single key 'translations'. "
            "The number of items must match the input exactly and the order must be preserved.\n\n"
            + json.dumps({"items": entries_payload}, ensure_ascii=False)
        )
        translated_text = self._strip_json_code_fence(
            self._request_translation(
                user_content,
                source_lang,
                target_lang,
                context=context,
                exception_rules=rules,
            )
        )
        try:
            response_json = json.loads(translated_text)
            entries = response_json.get("translations", [])
            if not isinstance(entries, list):
                raise ValueError("No 'translations' array in response")
        except (TypeError, ValueError, json.JSONDecodeError):
            logger.warning(
                "OpenAI slide response was not valid JSON; retrying with per-item requests."
            )
            results: dict[str, str] = {}
            for text in unique_texts:
                try:
                    results[text] = self._translate_single_text(
                        text,
                        source_lang,
                        target_lang,
                        context=context,
                        exception_rules=rules,
                    )
                except TranslationError as exc:
                    raise TranslationError(
                        f"OpenAI per-item translation failed for '{text[:60]}...': {exc}"
                    ) from exc
            return results

        results: dict[str, str] = {}
        for entry in entries:
            if not isinstance(entry, dict):
                continue
            item_id = str(entry.get("id", ""))
            item_text = entry.get("text")
            if item_id and isinstance(item_text, str):
                results[item_id] = self._restore_protected_replacements(item_text)
        return results
