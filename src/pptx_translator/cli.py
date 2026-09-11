"""Command-line interface for translating PowerPoint presentations."""

from __future__ import annotations

import argparse
import logging
import signal
import sys
from pathlib import Path

from .config import load_settings
from .exceptions_list import load_exception_rules
from .pptx_processor import PresentationTranslator
from .translators import create_translator


def _default_output_path(input_path: Path, target_lang: str) -> Path:
    suffix = input_path.suffix or ".pptx"
    return input_path.with_name(f"{input_path.stem}_{target_lang.lower()}{suffix}")


def _should_skip_translation(input_path: Path, target_lang: str) -> bool:
    suffix = f"_{target_lang.lower()}"
    return input_path.suffix.lower() == ".pptx" and input_path.stem.lower().endswith(suffix)


def _iter_pptx_files(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() == ".pptx" else []

    if recursive:
        iterator = input_path.rglob("*")
    else:
        iterator = input_path.iterdir()

    return sorted(
        {
            path
            for path in iterator
            if path.is_file() and path.suffix.lower() == ".pptx"
        },
        key=lambda p: str(p),
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-translate",
        description=(
            "Translates a PowerPoint presentation while preserving its "
            "original formatting. The source language is auto-detected."
        ),
    )
    parser.add_argument("input", type=Path, help="Path to the input .pptx file.")
    parser.add_argument(
        "-t",
        "--target",
        required=True,
        metavar="LANG",
        help="Target language code (ISO 639-1, e.g. 'en', 'fr', 'de').",
    )
    parser.add_argument(
        "-s",
        "--source",
        default=None,
        metavar="LANG",
        help=(
            "Source language code (ISO 639-1). If not provided, it is "
            "auto-detected from the presentation's content."
        ),
    )
    parser.add_argument(
        "-e",
        "--exceptions",
        type=Path,
        default=None,
        metavar="FILE",
        help=(
            "Path to a translation exceptions file. Each non-empty, "
            "non-comment line has the form 'source = destination' and is "
            "applied before translation; '{num}' can be used on both "
            "sides as a numeric wildcard (e.g. "
            "'Practica {num} = Practical Lesson {num}')."
        ),
    )
    parser.add_argument(
        "-p",
        "--provider",
        choices=("local", "openai"),
        default=None,
        help=(
            "Translation provider to use. Defaults to the value of "
            "TRANSLATOR_PROVIDER or 'local'."
        ),
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help=(
            "API key for the OpenAI-compatible provider. If omitted, the "
            "value from TRANSLATOR_API_KEY/OPENAI_API_KEY is used."
        ),
    )
    parser.add_argument(
        "--api-base-url",
        default=None,
        help=(
            "Base URL for the OpenAI-compatible API. Defaults to the value "
            "of TRANSLATOR_API_BASE_URL or https://api.openai.com/v1."
        ),
    )
    parser.add_argument(
        "--model",
        default=None,
        help=(
            "Model name to use with the OpenAI-compatible provider. Defaults "
            "to gpt-4o-mini."
        ),
    )
    parser.add_argument(
        "--request-delay",
        type=float,
        default=None,
        help=(
            "Seconds to wait between remote API requests. Defaults to "
            "TRANSLATOR_REQUEST_DELAY."
        ),
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help=(
            "When INPUT is a directory, also process .pptx files in nested "
            "subdirectories. Without this flag, only the files directly "
            "inside the directory are processed."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Preview the files that would be translated without writing any "
            "output files or modifying the filesystem."
        ),
    )
    parser.add_argument(
        "--remove-audio",
        action="store_true",
        help=(
            "Remove embedded audio objects from slides before translation. "
            "Disabled by default; can also be enabled via "
            "TRANSLATOR_REMOVE_AUDIO=true."
        ),
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help=(
            "Increase log verbosity, overriding TRANSLATOR_LOG_LEVEL. No "
            "flag: uses TRANSLATOR_LOG_LEVEL (ERROR/WARNING/INFO/DEBUG), or "
            "WARNING if unset; -v: INFO; -vv: DEBUG."
        ),
    )
    return parser


_LOG_LEVEL_NAMES = {
    "ERROR": logging.ERROR,
    "WARNING": logging.WARNING,
    "INFO": logging.INFO,
    "DEBUG": logging.DEBUG,
}


def _resolve_log_level(verbosity: int, configured_level: str | None) -> int:
    """Resolves the effective log level.

    Command-line ``-v``/``-vv`` flags always take precedence, since they
    are an explicit, per-run request. Otherwise, the ``TRANSLATOR_LOG_LEVEL``
    environment variable (``ERROR``/``WARNING``/``INFO``/``DEBUG``) is used if
    set to a recognized value, falling back to ``WARNING``.
    """

    if verbosity >= 2:
        return logging.DEBUG
    if verbosity == 1:
        return logging.INFO

    if configured_level:
        resolved = _LOG_LEVEL_NAMES.get(configured_level.strip().upper())
        if resolved is not None:
            return resolved

    return logging.WARNING


def _mask_secret(value: str | None) -> str:
    if not value:
        return "<unset>"
    prefix = value[:4]
    suffix = value[-2:] if len(value) > 2 else ""
    return f"{prefix}…{suffix}"


def _log_active_configuration(logger: logging.Logger, settings, provider_name: str) -> None:
    """Logs the selected provider and configuration in a readable form."""

    logger.info("Selected translation provider: %s", provider_name)
    logger.info("Audio removal enabled: %s", bool(settings.remove_audio))

    if not logger.isEnabledFor(logging.DEBUG):
        return

    logger.debug(
        "Active configuration: provider=%s, remove_audio=%s, max_chars_per_request=%s, request_delay_seconds=%s, max_retries=%s, retry_backoff_seconds=%s, api_base_url=%s, model=%s, log_level=%s, cache_dir=%s, api_key_configured=%s",
        provider_name,
        settings.remove_audio,
        settings.max_chars_per_request,
        settings.request_delay_seconds,
        settings.max_retries,
        settings.retry_backoff_seconds,
        settings.api_base_url,
        settings.model,
        settings.log_level,
        settings.cache_dir,
        bool(settings.api_key),
    )
    if settings.api_key:
        logger.debug("Configured API key: %s", _mask_secret(settings.api_key))


def _handle_sigint(signum: int, frame) -> None:
    """Translate Ctrl+C/SIGINT into a clean user-abort path."""
    raise KeyboardInterrupt(f"Received signal {signum}")


def main(argv: list[str] | None = None) -> int:
    try:
        if hasattr(signal, "SIGINT"):
            signal.signal(signal.SIGINT, _handle_sigint)

        parser = build_arg_parser()
        args = parser.parse_args(argv)

        settings = load_settings()
        log_level = _resolve_log_level(args.verbose, settings.log_level)

        logging.basicConfig(
            level=log_level,
            format="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%H:%M:%S",
            force=True,
        )
        logger = logging.getLogger("pptx_translator.cli")

        input_path: Path = args.input
        if not input_path.exists():
            logger.error("Input path does not exist: %s", input_path)
            return 1

        exception_rules = []
        if args.exceptions is not None:
            if not args.exceptions.exists():
                logger.error("Exceptions file does not exist: %s", args.exceptions)
                return 1
            try:
                exception_rules = load_exception_rules(args.exceptions)
            except ValueError as exc:
                logger.error(str(exc))
                return 1
            logger.info(
                "Loaded %d translation exception rule(s) from %s",
                len(exception_rules),
                args.exceptions,
            )

        try:
            translator = create_translator(
                settings,
                provider=args.provider,
                api_key=args.api_key,
                api_base_url=args.api_base_url,
                model=args.model,
                request_delay_seconds=args.request_delay,
            )
        except ValueError as exc:
            logger.error(str(exc))
            return 1

        _log_active_configuration(logger, settings, translator.name)

        if input_path.is_file():
            if input_path.suffix.lower() != ".pptx":
                logger.error("Input file must be a .pptx: %s", input_path)
                return 1
            files_to_process = [input_path]
        else:
            files_to_process = _iter_pptx_files(input_path, args.recursive)
            if not files_to_process:
                logger.error(
                    "No .pptx files found in directory '%s' (recursive=%s).",
                    input_path,
                    args.recursive,
                )
                return 1

        total_failed = 0
        for index, file_path in enumerate(files_to_process, start=1):
            if _should_skip_translation(file_path, args.target):
                logger.warning(
                    "Skipping file '%s': it already ends with the target-language suffix '_%s'.",
                    file_path,
                    args.target.lower(),
                )
                continue

            output_path = _default_output_path(file_path, args.target)

            logger.info(
                "[%d/%d] Translation provider: %s | Input: %s | Output: %s | Target: %s",
                index,
                len(files_to_process),
                translator.name,
                file_path,
                output_path,
                args.target,
            )

            if args.dry_run:
                if output_path.exists():
                    logger.warning(
                        "Dry run: file '%s' already exists and would be overwritten.",
                        output_path,
                    )
                else:
                    logger.warning(
                        "Dry run: translation would be written to '%s'.",
                        output_path,
                    )
                continue

            output_path.parent.mkdir(parents=True, exist_ok=True)

            presentation_translator = PresentationTranslator(
                translator,
                exception_rules=exception_rules,
                remove_audio=args.remove_audio or settings.remove_audio,
            )
            try:
                stats = presentation_translator.translate(
                    str(file_path),
                    str(output_path),
                    target_lang=args.target,
                    source_lang=args.source,
                )
            except Exception as exc:  # noqa: BLE001
                logger.exception("Error translating the presentation '%s': %s", file_path, exc)
                total_failed += 1
                continue

            logger.info("Detected/used source language: %s", stats.detected_source_lang)
            logger.info("Slides processed: %d", stats.slides)
            logger.info("Audio elements removed: %d", stats.audio_removed)
            logger.info("Text boxes treated as figures: %d", stats.figures_detected)
            logger.info(
                "Unique texts translated: %d (total units rewritten: %d)",
                stats.unique_texts,
                stats.translated_units,
            )
            if stats.failed_texts:
                logger.warning(
                    "%d text(s) could not be translated and were left in the original language.",
                    stats.failed_texts,
                )
            if log_level <= logging.DEBUG:
                for reason in stats.reasons_log:
                    logger.debug("Figure heuristic: %s", reason)
            logger.info("Translated presentation saved to: %s", output_path)

        if len(files_to_process) > 1:
            if args.dry_run:
                logger.info(
                    "Batch dry run complete: %d file(s) would be processed, %d error(s).",
                    len(files_to_process),
                    total_failed,
                )
            else:
                logger.info(
                    "Batch complete: %d file(s) processed, %d error(s).",
                    len(files_to_process),
                    total_failed,
                )

        return 1 if total_failed == len(files_to_process) else 0
    except KeyboardInterrupt:
        print("Execution aborted by the user", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
