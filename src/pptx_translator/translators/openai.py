"""HTTP-based translation providers compatible with OpenAI-style APIs."""

from __future__ import annotations

import logging
import os
from typing import Any

import requests

from .base import BaseTranslator, TranslationError

logger = logging.getLogger(__name__)


class OpenAITranslator(BaseTranslator):
    """Translator backed by an OpenAI-compatible HTTP endpoint."""

    name = "openai"

    def __init__(
        self,
        api_key: str | None = None,
        api_base_url: str | None = None,
        model: str | None = None,
        *args,
        **kwargs,
    ) -> None:
        self.api_key = (api_key or os.getenv("TRANSLATOR_API_KEY") or os.getenv("OPENAI_API_KEY") or "").strip()
        self.api_base_url = (api_base_url or os.getenv("TRANSLATOR_API_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
        self.model = model or os.getenv("TRANSLATOR_MODEL") or "gpt-4o-mini"
        if not self.api_key:
            raise ValueError(
                "An API key is required for remote translation providers. "
                "Set TRANSLATOR_API_KEY or pass api_key=... ."
            )
        super().__init__(*args, **kwargs)

    def on_presentation_start(self, context: str | None = None) -> None:
        logger.info(
            "Using first-slide presentation context for OpenAI-compatible translation: %s",
            context,
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

    def _build_system_content(self, context: str | None = None) -> str:
        base = (
            "You are a specialist technical translator for PowerPoint slides and presentation materials. "
            "Translate accurately and idiomatically for the target language, while preserving meaning, technical conventions, and the exact structure of fixed labels and protected terms. "
            "Return only plain translated text."
        )
        if not context:
            return base

        safe_context = " ".join(context.replace("\r", " ").replace("\n", " ").split())
        return (
            f"Presentation context: {safe_context}. "
            "Use this context to preserve the correct domain terminology, title conventions, author names, affiliations, and presentation-specific phrasing. "
            + base
        )

    def _translate_one(
        self,
        text: str,
        source_lang: str,
        target_lang: str,
        context: str | None = None,
    ) -> str:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        prompt = (
            f"Translate the following {source_lang} text into {target_lang}. "
            "The text comes from a PowerPoint presentation and may be a slide sentence, a short label, "
            "a diagram caption, or part of a speaker note.\n\n"
            "Requirements:\n"
            "1. Return only the final translated text, with no explanation, no markdown, no code fences, and no extra commentary.\n"
            "2. Preserve the original meaning, tone, and intent exactly; do not add or remove information.\n"
            "3. Use natural, idiomatic target-language wording for the context, especially for short labels and technical terminology.\n"
            "4. Preserve technical terms, acronyms, product names, units, numbers, dates, labels, and placeholders exactly when they are fixed domain terms or protected identifiers.\n"
            "5. Keep punctuation, capitalization, and sentence boundaries consistent with the source.\n"
            "6. If the source is a short label or caption, translate it as a concise label, not as a long explanatory sentence.\n"
            "7. If the text is already in the target language or is a proper noun, leave it unchanged unless the target language convention explicitly requires a different form.\n"
            "8. Do not hallucinate or invent content that is not present in the source text.\n\n"
            f"Text to translate:\n{text}"
        )
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": self._build_system_content(context),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
        }
        response = requests.post(
            f"{self.api_base_url}/chat/completions",
            headers=headers,
            json=payload,
            timeout=60,
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
