"""HTTP-based translation providers compatible with OpenAI-style APIs."""

from __future__ import annotations

import json
import logging
import os
import re
from typing import Any

import requests

from .base import BaseTranslator, TranslationError

logger = logging.getLogger(__name__)


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

    def _build_system_content(
        self,
        context: str | None = None,
        protected_replacements: dict[str, str] | None = None,
    ) -> str:
        protected_replacements = protected_replacements or getattr(self, "_protected_replacements", {})
        replacement_text = ""
        if protected_replacements:
            protected_entries = [
                f"{token} -> {value}"
                for token, value in sorted(
                    protected_replacements.items(),
                    key=lambda item: len(item[0]),
                    reverse=True,
                )
            ]
            replacement_text = (
                "Protected identifiers that must remain exactly unchanged and must be restored verbatim are: "
                + "; ".join(protected_entries)
                + ". "
            )

        protected_guidance = (
            "Protected placeholders are not ordinary words: they are internal markers that must remain exactly unchanged, including their casing, spacing, and punctuation. "
            "Do not translate, split, expand, paraphrase, or alter any placeholder token such as zqkpptx...vxq, and do not insert any extra spaces around them. "
            "Treat them as immutable fixed identifiers. "
        )
        base = (
            "You are a specialist technical translator for PowerPoint slides and presentation materials. "
            "Translate accurately and idiomatically for the target language, while preserving meaning, technical conventions, and the exact structure of fixed labels and protected terms. "
            "Return only plain translated text. "
            + replacement_text
            + protected_guidance
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

    def _request_translation(self, user_content: str, source_lang: str, target_lang: str, context: str | None = None) -> str:
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
                        context,
                        getattr(self, "_protected_replacements", {}),
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
        logger.debug("OpenAI request payload: %s", json.dumps(payload, ensure_ascii=False, indent=2))

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
    ) -> str:
        user_content = (
            f"Translate the following text from {source_lang} to {target_lang}. "
            "Return only the translated text and nothing else.\n\n"
            f"{text}"
        )
        translated_text = self._request_translation(user_content, source_lang, target_lang, context=context)
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
    ) -> dict[str, str]:
        """Translate every text fragment from a slide in a single API call."""

        unique_texts = list(dict.fromkeys(texts))
        if not unique_texts:
            return {}

        user_content = (
            "Translate each item in the JSON array below from "
            f"{source_lang} to {target_lang}. Return a JSON object with a "
            "single key 'translations' whose value is a list of objects like "
            '{"id": "source text", "text": "translated text"}. Keep the same order as the input array. Do not include any extra text outside the JSON.\n\n'
            + json.dumps({
                "items": [{"id": text, "text": text} for text in unique_texts]
            })
        )
        translated_text = self._request_translation(user_content, source_lang, target_lang, context=context)
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
                    results[text] = self._translate_single_text(text, source_lang, target_lang, context=context)
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
