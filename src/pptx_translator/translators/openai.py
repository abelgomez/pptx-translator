"""HTTP-based translation providers compatible with OpenAI-style APIs."""

from __future__ import annotations

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
        *args,
        **kwargs,
    ) -> None:
        self.api_key = (api_key or os.getenv("TRANSLATOR_API_KEY") or os.getenv("OPENAI_API_KEY") or "").strip()
        self.api_base_url = (api_base_url or os.getenv("TRANSLATOR_API_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
        self.model = model or os.getenv("TRANSLATOR_MODEL") or "gpt-4o-mini"
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
                    "content": self._build_system_content(
                        context,
                        getattr(self, "_protected_replacements", {}),
                    ),
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

        translated = self._extract_text(data)
        translated = self._restore_protected_replacements(translated)
        protected_token_pattern = re.compile(r"zqkpptx[a-z]+vxq", re.IGNORECASE)
        if protected_token_pattern.search(translated):
            logger.warning(
                "OpenAI response still contains protected exception placeholders; restoring protected fragments after translation."
            )
        return translated
