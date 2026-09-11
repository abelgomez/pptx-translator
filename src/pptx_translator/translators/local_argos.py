"""Local (offline) translation provider using Argos Translate.

Argos Translate is an open-source neural machine translation engine that
runs entirely locally (no internet connection needed after the language
models are downloaded the first time), which means:

* It does not consume any quota or limit from an external API.
* It cannot be blocked due to request abuse.
* Its translations have reasonable quality for most supported language
  pairs (based on OpenNMT/CTranslate2 models trained on OPUS data).

This is the default provider of this application. The first time a
particular language pair is translated, the corresponding language package
is automatically downloaded (once, then cached on disk) from the official
Argos Translate repository.
"""

from __future__ import annotations

import logging
import threading

from .base import BaseTranslator, TranslationError

logger = logging.getLogger(__name__)

# Installing/downloading Argos Translate packages is not thread-safe and
# there is no point in parallelizing it; it is serialized with a process lock.
_install_lock = threading.Lock()
_installed_pairs: set[tuple[str, str]] = set()


def _sync_argos_internal_logging() -> None:
    """Keep Argos Translate's own internal logging quiet unless the app is
    running in DEBUG mode.

    ``argostranslate.utils`` unconditionally sets its logger to ``INFO``
    (or ``DEBUG`` if the library itself is in debug mode) at import time.
    Because that is an *explicit* logger-level override, it can bypass the
    application's root logger configuration and leak noisy internal output
    even when the project is configured to show only warnings/errors.

    To avoid that, we force Argos' internal loggers to ``DEBUG`` only when
    the application is globally set to DEBUG, and otherwise silently mute
    them (``CRITICAL`` so that INFORMATION/WARNING/ERROR records never pass
    through).
    """

    root_level = logging.getLogger().getEffectiveLevel()
    required_level = logging.DEBUG if root_level <= logging.DEBUG else logging.CRITICAL

    for noisy_logger_name in ("argostranslate", "argostranslate.utils"):
        logging.getLogger(noisy_logger_name).setLevel(required_level)

    if required_level > logging.DEBUG:
        logging.getLogger("stanza").disabled = True

class LocalArgosTranslator(BaseTranslator):
    """Local translator based on Argos Translate (no network calls)."""

    name = "local"

    def __init__(self, *args, auto_install: bool = True, **kwargs) -> None:
        # There is no point in throttling requests to a local engine: it
        # runs on the same machine and there is no risk of being blocked.
        kwargs.setdefault("request_delay_seconds", 0.0)
        super().__init__(*args, **kwargs)
        self.auto_install = auto_install

    def _ensure_language_pair(self, source_lang: str, target_lang: str) -> None:
        pair = (source_lang, target_lang)
        if pair in _installed_pairs:
            return

        import argostranslate.package as argos_package
        import argostranslate.translate as argos_translate

        _sync_argos_internal_logging()

        with _install_lock:
            if pair in _installed_pairs:
                return

            installed_languages = argos_translate.get_installed_languages()
            from_lang = next(
                (lang for lang in installed_languages if lang.code == source_lang), None
            )
            if from_lang is not None:
                has_direct_translation = any(
                    translation.to_lang.code == target_lang
                    for translation in from_lang.translations_from
                )
                if has_direct_translation:
                    _installed_pairs.add(pair)
                    return

            if not self.auto_install:
                raise TranslationError(
                    f"No local model is installed for '{source_lang}' -> "
                    f"'{target_lang}' and automatic installation is disabled. "
                    "Install it manually or enable 'auto_install'."
                )

            logger.info(
                "Downloading the local translation model %s -> %s "
                "(this only happens once; it may take a few minutes)...",
                source_lang,
                target_lang,
            )
            argos_package.update_package_index()
            available_packages = argos_package.get_available_packages()
            package = next(
                (
                    pkg
                    for pkg in available_packages
                    if pkg.from_code == source_lang and pkg.to_code == target_lang
                ),
                None,
            )
            if package is None:
                raise TranslationError(
                    f"Argos Translate does not offer a direct model for "
                    f"'{source_lang}' -> '{target_lang}'. Try another "
                    "provider or translate through an intermediate language "
                    "(e.g. English)."
                )
            downloaded_path = package.download()
            argos_package.install_from_path(downloaded_path)
            _installed_pairs.add(pair)

    def _translate_slide(
        self,
        texts: list[str],
        source_lang: str,
        target_lang: str,
        context: str | None = None,
    ) -> dict[str, str]:
        import argostranslate.translate as argos_translate

        _sync_argos_internal_logging()

        source = source_lang if source_lang not in ("", "auto") else "es"
        self._ensure_language_pair(source, target_lang)
        return {text: argos_translate.translate(text, source, target_lang) for text in dict.fromkeys(texts)}
