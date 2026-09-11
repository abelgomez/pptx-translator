"""User-configurable translation exceptions.

Some fragments of text must not be translated literally (e.g. a fixed
label that has a specific, non-literal equivalent in the target language,
or an acronym that must always be kept unchanged). Instead of hard-coding
such rules in the source code, this module loads them from a plain-text
file supplied by the user (one rule per line) and applies them either
before or around the translation of every text unit.

File format
-----------
Each non-empty, non-comment line has the form::

    source expression = destination expression

Lines starting with ``#`` are treated as comments and ignored, as are
blank lines. Both sides are trimmed of surrounding whitespace.

An optional single-character prefix selects how the rule is applied:

* ``<`` -- **Pre-translation substitution.** The source expression is
  replaced by the destination expression *before* the text is sent to the
  translator, and the result is translated normally together with the rest
  of the sentence. Use this when the destination expression is already
  written in the target language and you are fine with the translator
  possibly still adjusting it grammatically while translating the
  surrounding text.
* ``!`` -- **Strict, case-sensitive.** The source expression is matched
  exactly (case-sensitive) against the original text, protected so the
  translator cannot alter it, and the destination expression is
  guaranteed to appear verbatim in the final output. Use this for
  acronyms/identifiers that must never change case or spelling (e.g.
  ``SGA``).
* ``~`` -- **Strict, case-insensitive.** Same guarantee as ``!``, but the
  source expression is matched ignoring case.

If no prefix is given, the rule behaves like ``~`` (strict,
case-insensitive), which matches the historical default behavior of this
tool.

The placeholder ``{num}`` acts as a wildcard that matches one or more
digits:

* On the left-hand side (the source expression), ``{num}`` matches any
  run of digits found in the original text.
* On the right-hand side (the destination expression), each ``{num}`` is
  replaced, in order, with the digits captured by the corresponding
  ``{num}`` wildcard on the left-hand side.

Example::

    ~Práctica {num} = Practical Lesson {num}
    !SGA = SGA

The first rule turns "Práctica 1" into "Practical Lesson 1" (protected
from mistranslation, matched regardless of case). The second guarantees
the acronym "SGA" is never altered by the translator (e.g. turned into
"USG" or similar).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Sequence

_WILDCARD_TOKEN = "{num}"
_WILDCARD_PATTERN = r"(\d+)"

# Placeholders used to protect "strict"/"case-insensitive" exceptions while
# the surrounding text is being translated. They are built purely out of
# lowercase letters (no digits, underscores, or other punctuation) so they
# form a single unbroken "word": neural translation models and their
# detokenizers are much more likely to leave such a token untouched, since
# punctuation characters (like the underscores used by a previous version
# of this module) are frequently treated as separate tokens and can end up
# with spaces inserted around them by the detokenizer, corrupting the
# placeholder beyond recognition (e.g. "__PPTX_TRANSLATOR_EXC_0__" turning
# into "_ _ PPTX _ TRANSLATOR _ EXC _ 0 _ _").
_TOKEN_PREFIX = "zqkpptx"
_TOKEN_SUFFIX = "vxq"
_TOKEN_LETTERS = "abcdefghijklmnopqrstuvwxyz"


class ExceptionMode(Enum):
    """How a single exception rule is applied."""

    PRE_TRANSLATION = "pre"  # '<' prefix: plain substitution before translating
    STRICT = "strict"  # '!' prefix: case-sensitive protect + restore
    CASE_INSENSITIVE = "ci"  # '~' prefix (or no prefix): case-insensitive protect + restore


_MODE_PREFIXES = {
    "<": ExceptionMode.PRE_TRANSLATION,
    "!": ExceptionMode.STRICT,
    "~": ExceptionMode.CASE_INSENSITIVE,
}


@dataclass
class ExceptionRule:
    """A single compiled "source = destination" exception rule."""

    mode: ExceptionMode
    pattern: re.Pattern[str]
    replacement_template: str
    raw_source: str
    raw_destination: str


def _index_to_token_id(index: int) -> str:
    """Converts an integer index into a purely alphabetic, base-26 id
    (0 -> "a", 1 -> "b", ..., 25 -> "z", 26 -> "aa", ...), the same scheme
    spreadsheets use for column names. Using only letters keeps the
    generated placeholder as a single unbroken word (see module docstring).
    """

    n = index + 1
    letters: list[str] = []
    while n > 0:
        n, remainder = divmod(n - 1, len(_TOKEN_LETTERS))
        letters.append(_TOKEN_LETTERS[remainder])
    return "".join(reversed(letters))


def _make_placeholder_token(index: int) -> str:
    return f"{_TOKEN_PREFIX}{_index_to_token_id(index)}{_TOKEN_SUFFIX}"


def _compile_source_pattern(source_expr: str, case_sensitive: bool) -> re.Pattern[str]:
    """Builds a regex from a source expression.

    Every literal part of the expression is escaped so it is matched
    verbatim; every ``{num}`` placeholder becomes a capturing group that
    matches one or more digits.
    """

    segments = source_expr.split(_WILDCARD_TOKEN)
    escaped_segments = [re.escape(segment) for segment in segments]
    pattern_text = _WILDCARD_PATTERN.join(escaped_segments)
    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile(pattern_text, flags=flags)


def _build_replacement(template: str, match: re.Match[str]) -> str:
    """Fills the ``{num}`` placeholders in the destination template with
    the digits captured by the source pattern, in order."""

    groups = match.groups()
    segments = template.split(_WILDCARD_TOKEN)
    result_parts: list[str] = [segments[0]]
    for index, segment in enumerate(segments[1:]):
        captured = groups[index] if index < len(groups) else _WILDCARD_TOKEN
        result_parts.append(captured)
        result_parts.append(segment)
    return "".join(result_parts)


def load_exception_rules(path: str | Path) -> list[ExceptionRule]:
    """Parses an exceptions file into a list of :class:`ExceptionRule`.

    Raises ``ValueError`` if a non-empty, non-comment line does not follow
    the ``[<|!|~]source expression = destination expression`` format.
    """

    rules: list[ExceptionRule] = []
    file_path = Path(path)
    with file_path.open("r", encoding="utf-8") as handle:
        for line_number, raw_line in enumerate(handle, start=1):
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue

            mode = ExceptionMode.CASE_INSENSITIVE
            if line[0] in _MODE_PREFIXES:
                mode = _MODE_PREFIXES[line[0]]
                line = line[1:].strip()

            if "=" not in line:
                raise ValueError(
                    f"Invalid exception rule at {file_path}:{line_number}: "
                    f"{raw_line.strip()!r}. Expected format "
                    "'[<|!|~]source expression = destination expression'."
                )
            source_expr, destination_expr = line.split("=", 1)
            source_expr = source_expr.strip()
            destination_expr = destination_expr.strip()
            if not source_expr:
                raise ValueError(
                    f"Invalid exception rule at {file_path}:{line_number}: "
                    "the source expression cannot be empty."
                )
            case_sensitive = mode is ExceptionMode.STRICT
            rules.append(
                ExceptionRule(
                    mode=mode,
                    pattern=_compile_source_pattern(source_expr, case_sensitive),
                    replacement_template=destination_expr,
                    raw_source=source_expr,
                    raw_destination=destination_expr,
                )
            )
    return rules


def split_rules_by_mode(
    rules: Sequence[ExceptionRule],
) -> tuple[list[ExceptionRule], list[ExceptionRule]]:
    """Splits rules into ``(pre_translation_rules, protected_rules)``.

    ``pre_translation_rules`` (``<``) must be applied with
    :func:`apply_exception_rules` before translating. ``protected_rules``
    (``!``/``~``) must be applied with :func:`mask_exception_rules` before
    translating and :func:`restore_exception_rules` afterward.
    """

    pre_rules = [r for r in rules if r.mode is ExceptionMode.PRE_TRANSLATION]
    protected_rules = [r for r in rules if r.mode is not ExceptionMode.PRE_TRANSLATION]
    return pre_rules, protected_rules


def apply_exception_rules(text: str, rules: Sequence[ExceptionRule]) -> str:
    """Applies every exception rule, in order, to ``text``.

    This performs a plain substitution with no protection guarantee: the
    resulting text is translated normally afterward, so the translator may
    still adjust the substituted fragment (used for ``<`` "pre-translation"
    rules).
    """

    for rule in rules:
        text = rule.pattern.sub(
            lambda match, template=rule.replacement_template: _build_replacement(
                template, match
            ),
            text,
        )
    return text


def mask_exception_rules(
    text: str,
    rules: Sequence[ExceptionRule],
    start_index: int = 0,
) -> tuple[str, dict[str, str]]:
    """Replaces every matching exception with a protected placeholder token.

    This prevents the translator from modifying the protected term before
    it is re-inserted (see :func:`restore_exception_rules`) once
    translation is complete. Only ``!``/``~`` ("strict") rules should be
    passed here; ``<`` rules must go through :func:`apply_exception_rules`
    instead.

    ``start_index`` allows callers to reserve a globally unique namespace of
    placeholder ids so that distinct text fragments never reuse the same
    placeholder token (which would otherwise cause exception values to be
    mixed together across different translated texts).
    """

    protected = text
    replacements: dict[str, str] = {}
    token_index = start_index

    for rule in rules:
        def _replace(match: re.Match[str], template: str = rule.replacement_template) -> str:
            nonlocal token_index
            token = _make_placeholder_token(token_index)
            token_index += 1
            replacements[token] = _build_replacement(template, match)
            return token

        protected = rule.pattern.sub(_replace, protected)

    return protected, replacements


def restore_exception_rules(text: str, replacements: dict[str, str]) -> str:
    """Replaces protected exception tokens with their final destination text.

    Matching is done case-insensitively because translation engines
    sometimes change the casing of a token (e.g. capitalizing the first
    letter of a sentence), even when they otherwise leave it intact.
    """

    restored = text
    for token, value in sorted(replacements.items(), key=lambda item: len(item[0]), reverse=True):
        restored = re.sub(re.escape(token), lambda _m, v=value: v, restored, flags=re.IGNORECASE)
    return restored


def placeholder_token_pattern() -> str:
    """Returns a human-readable pattern describing generated placeholder
    tokens (e.g. for inclusion in a translation-provider system prompt so
    it knows to leave them untouched)."""

    return f"{_TOKEN_PREFIX}...{_TOKEN_SUFFIX}"
