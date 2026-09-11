"""Automatic detection of the source language of a presentation."""

from __future__ import annotations

import logging

from langdetect import DetectorFactory, LangDetectException, detect

# Makes detection deterministic (langdetect uses an internal pseudo-random
# generator for n-gram sampling).
DetectorFactory.seed = 0

logger = logging.getLogger(__name__)


def detect_language(sample_text: str, fallback: str = "es") -> str:
    """Detects the language (ISO 639-1 code, e.g. ``es``, ``en``) of a text.

    An aggregated sample of text from several slides is used to give
    langdetect enough context. If detection fails (empty text, too short,
    symbols only, etc.) ``fallback`` is returned instead.
    """

    text = sample_text.strip()
    if not text:
        logger.warning(
            "Not enough text was found to detect the source language; "
            "falling back to '%s'.",
            fallback,
        )
        return fallback

    try:
        code = detect(text)
    except LangDetectException:
        logger.warning(
            "The source language could not be auto-detected; falling back to '%s'.",
            fallback,
        )
        return fallback

    # langdetect sometimes returns regional variants (zh-cn, etc.); we keep
    # only the 2-letter language code when possible.
    return code.split("-")[0]
