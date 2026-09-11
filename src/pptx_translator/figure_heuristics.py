"""Heuristics to decide whether a text box is part of a figure.

When a text box belongs to a figure (for example, a label inside a diagram
made up of shapes, arrows, ovals, etc.), the line breaks it contains are not
real paragraph separators, but were manually inserted so the text fits
visually inside the shape. Therefore, all of its content must be treated as
a single piece of text before translating it, rather than as independent
paragraphs.

There are two ways to detect this situation:

1. The text box is grouped (``GROUP``) together with other shapes: it is
   always assumed to be part of a figure.
2. The text box is NOT grouped, but based on its appearance (few words per
   line, short total text, small label-like geometry, absence of sentence-
   ending punctuation) it is very likely a standalone figure label.

The geometry is important: if the box is small and contains little text,
then the decisive signal is whether it looks like a label (roughly square,
more tall than wide, or narrow like a one-word label), not just whether it
contains line breaks. If the text is substantial (many words / long lines),
its line breaks are preserved because they are likely real layout content,
not figure-internal formatting.
"""

from __future__ import annotations

from dataclasses import dataclass

# Heuristic thresholds (tunable if needed).
MAX_WORDS_PER_LINE_AVG = 3.5
MAX_TOTAL_WORDS = 12
MAX_AREA_RATIO = 0.06  # shape area / slide area
MAX_NARROW_WIDTH_RATIO = 0.12  # width relative to slide width: one-word-like label
SENTENCE_END_CHARS = (".", "!", "?", ";", ":")


@dataclass
class FigureHeuristicResult:
    is_figure: bool
    reasons: list[str]


def _text_lines(text: str) -> list[str]:
    return [line for line in text.split("\n") if line.strip()]


def looks_like_figure_label(
    text: str,
    shape_width_emu: int | None,
    shape_height_emu: int | None,
    slide_width_emu: int | None,
    slide_height_emu: int | None,
) -> FigureHeuristicResult:
    """Applies the heuristic to a NON-grouped text box.

    A label-like box is only treated as a figure if the text is short and the
    box geometry resembles a label: roughly square, taller than wide, or as
    narrow as a one-word label. If the text is substantial, its line breaks
    are preserved because they are much more likely to be meaningful layout
    or paragraph separators.
    """

    lines = _text_lines(text)
    reasons: list[str] = []

    if not lines:
        return FigureHeuristicResult(False, reasons)

    # If the text does not even contain manual line breaks, the question of
    # "joining lines" does not apply: there is nothing to join.
    if len(lines) < 2:
        return FigureHeuristicResult(False, reasons)

    words_per_line = [len(line.split()) for line in lines]
    avg_words_per_line = sum(words_per_line) / len(words_per_line)
    total_words = sum(words_per_line)

    if avg_words_per_line > MAX_WORDS_PER_LINE_AVG:
        reasons.append(f"long lines (avg={avg_words_per_line:.1f})")
        return FigureHeuristicResult(False, reasons)

    if total_words > MAX_TOTAL_WORDS:
        reasons.append(f"too much text for a figure label ({total_words} words)")
        return FigureHeuristicResult(False, reasons)

    if avg_words_per_line <= MAX_WORDS_PER_LINE_AVG:
        reasons.append(f"few words per line (avg={avg_words_per_line:.1f})")

    if total_words <= MAX_TOTAL_WORDS:
        reasons.append(f"few total words ({total_words})")

    geometry_reasons: list[str] = []
    is_squareish = False
    is_tall = False
    is_word_like_width = False

    if (
        shape_width_emu is not None
        and shape_height_emu is not None
        and slide_width_emu is not None
        and slide_height_emu is not None
    ):
        shape_area = shape_width_emu * shape_height_emu
        slide_area = slide_width_emu * slide_height_emu
        if slide_area > 0:
            ratio = shape_area / slide_area
            if ratio <= MAX_AREA_RATIO:
                geometry_reasons.append(f"box is small relative to the slide ({ratio:.1%})")

        if shape_height_emu > 0 and shape_width_emu > 0:
            aspect_ratio = shape_width_emu / shape_height_emu
            if 0.75 <= aspect_ratio <= 1.35:
                is_squareish = True
                geometry_reasons.append("shape is approximately square")
            if shape_height_emu > shape_width_emu * 1.2:
                is_tall = True
                geometry_reasons.append("shape is taller than it is wide")

        if slide_width_emu > 0 and shape_width_emu / slide_width_emu <= MAX_NARROW_WIDTH_RATIO:
            is_word_like_width = True
            geometry_reasons.append(
                f"shape width is about one word wide ({shape_width_emu / slide_width_emu:.1%})"
            )

    # Absence of sentence-ending punctuation at the end of each line:
    # a sign that the lines are not independent complete sentences.
    lines_without_sentence_end = sum(
        1 for line in lines if not line.strip().endswith(SENTENCE_END_CHARS)
    )
    if lines_without_sentence_end == len(lines):
        reasons.append("no line ends with sentence-closing punctuation")

    if geometry_reasons:
        reasons.extend(geometry_reasons)

    geometry_score = sum(
        1
        for signal in (is_squareish, is_tall, is_word_like_width)
        if signal
    )
    if geometry_score >= 1 and len(reasons) >= 3:
        return FigureHeuristicResult(True, reasons)

    return FigureHeuristicResult(False, reasons)
