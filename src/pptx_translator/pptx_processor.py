"""Core logic for translating PowerPoint presentations.

This module walks through every slide of a presentation, locates the
translatable text (respecting the paragraph/bullet structure of normal
content and joining the lines of figures), removes embedded audio, and
writes the translated text back while preserving the original formatting as
much as possible (font, size, color, alignment, bullets).
"""

from __future__ import annotations

import copy
import logging
import re
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from pptx.slide import Slide
from pptx.util import Length

from .exceptions_list import (
    ExceptionRule,
    apply_exception_rules,
    mask_exception_rules,
    restore_exception_rules,
    split_rules_by_mode,
)
from .figure_heuristics import looks_like_figure_label
from .lang import detect_language
from .media import remove_audio_from_slide
from .translators.base import BaseTranslator

logger = logging.getLogger(__name__)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WHITESPACE_RE = re.compile(r"\s+")


@dataclass
class TranslationJob:
    """A pending unit of work to translate.

    ``kind`` is either ``"paragraph"`` (translates and rewrites a single
    paragraph, preserving its structure), ``"paragraph_run"`` (translates
    a single styled run inside a paragraph while keeping that run's own
    formatting), or ``"figure"`` (translates the full text of a text box
    and rewrites it as a single paragraph, automatically resizing the text
    to fit the shape).
    """

    kind: str
    target: object  # _Paragraph for "paragraph"/"paragraph_run", TextFrame for "figure"
    original_text: str
    text_frame: object | None = None
    slide_index: int = 0
    exception_replacements: dict[str, str] = field(default_factory=dict)
    run_segments: list[tuple[object, str, str, str]] = field(default_factory=list)


@dataclass
class TranslationStats:
    slides: int = 0
    audio_removed: int = 0
    unique_texts: int = 0
    translated_units: int = 0
    figures_detected: int = 0
    detected_source_lang: str | None = None
    failed_texts: int = 0
    reasons_log: list[str] = field(default_factory=list)


def _clean_joined_text(text: str) -> str:
    """Collapses line breaks/vertical tabs into single spaces."""

    text = text.replace("\v", " ").replace("\x0b", " ")
    text = _WHITESPACE_RE.sub(" ", text)
    return text.strip()


def _paragraph_plain_text(paragraph) -> str:
    return _clean_joined_text(paragraph.text)


def _paragraph_has_text_run(paragraph) -> bool:
    """True if the paragraph contains at least one ``<a:r>`` (normal text
    run), as opposed to only containing dynamic fields (``<a:fld>``, e.g.
    the slide number or the date) which must not be translated.
    """

    return paragraph._p.find(qn("a:r")) is not None  # noqa: SLF001


def _iter_shapes_recursive(shapes, in_group: bool):
    """Yields (shape, in_group) tuples, also walking into groups."""

    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes_recursive(shape.shapes, True)
            continue
        yield shape, in_group


def _set_paragraph_text(paragraph, new_text: str, force: bool = False) -> None:
    """Rewrites a paragraph with a single piece of text.

    Keeps the formatting (rPr) of the first existing run and the paragraph
    properties (pPr: bullets, alignment, indentation...). Removes any
    manual line break (``<a:br/>``) and additional runs, since all content
    is now represented in a single run.

    If ``force`` is ``False`` (the default) and the paragraph only contains
    a dynamic field (``<a:fld/>``, e.g. slide number or date), it is left
    untouched, since that content is recalculated automatically by
    PowerPoint. ``force=True`` is used when merging the full text of a
    figure, where any remaining content should indeed be overwritten.
    """

    p_elem = paragraph._p  # noqa: SLF001
    run_elements = p_elem.findall(qn("a:r"))

    if run_elements:
        keep_run = run_elements[0]
        t_elem = keep_run.find(qn("a:t"))
        if t_elem is None:
            t_elem = keep_run.makeelement(qn("a:t"), {})
            keep_run.append(t_elem)
        t_elem.text = new_text
        for child in list(p_elem):
            if child is keep_run or child.tag == qn("a:pPr"):
                continue
            p_elem.remove(child)
        return

    # There were no normal text runs (the paragraph only contained a dynamic
    # field -<a:fld/>-, such as the slide number or the date, or manual line
    # breaks). Dynamic fields are recalculated by PowerPoint automatically,
    # so they are left untouched unless force=True.
    has_field = p_elem.find(qn("a:fld")) is not None
    if has_field and not force:
        return

    for child in list(p_elem):
        if child.tag in (qn("a:br"), qn("a:fld")):
            p_elem.remove(child)
    if new_text:
        run = paragraph.add_run()
        run.text = new_text


def _set_paragraph_runs(paragraph, run_segments: list[tuple[object, str]], translated_texts: list[str]) -> None:
    """Rebuilds a paragraph from its original styled runs.

    This keeps run-level formatting intact, which matters for speaker notes
    where only a short lead-in phrase (for example, a bold concept name) is
    emphasized while the rest of the sentence remains in the normal style.
    """

    if not run_segments:
        return

    p_elem = paragraph._p  # noqa: SLF001
    for child in list(p_elem):
        if child.tag in (qn("a:r"), qn("a:fld"), qn("a:br")):
            p_elem.remove(child)

    for original_run, translated_text in zip(run_segments, translated_texts):
        run = paragraph.add_run()
        run.text = translated_text

        run_elem = run._r  # noqa: SLF001
        original_r_pr = original_run._r.find(qn("a:rPr"))  # noqa: SLF001
        if original_r_pr is not None:
            existing_r_pr = run_elem.find(qn("a:rPr"))
            if existing_r_pr is not None:
                run_elem.remove(existing_r_pr)
            run_elem.insert(0, copy.deepcopy(original_r_pr))


def _find_template_run_properties(text_frame):
    """Finds the first available ``<a:rPr>`` in the text box."""

    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            r_pr = run._r.find(qn("a:rPr"))  # noqa: SLF001
            if r_pr is not None:
                return copy.deepcopy(r_pr)
    return None


def _apply_text_frame_fit(text_frame, new_text: str) -> None:
    """Enables automatic wrapping/fit for translated text boxes."""
    if not new_text or text_frame is None:
        return

    # Auto-fit the text to the shape only when the translated text has
    # more than one word: a single word is usually a short label/acronym,
    # and forcing it to shrink to fit would make it unreadable, so it is
    # left as-is even if it overflows the shape.
    word_count = len(new_text.split())
    if word_count <= 1:
        return

    text_frame.word_wrap = True
    try:
        text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:  # pragma: no cover - some shapes don't support auto_size
        logger.debug("Could not set auto_size on a translated text box.")


def _set_text_frame_single_text(text_frame, new_text: str) -> None:
    """Replaces all the content of a text box with a single paragraph.

    Used for figures: the text of all the original lines/paragraphs has
    already been translated as a single string, so all that's left is to
    leave a single paragraph with that string, reusing the original
    formatting, and enabling automatic text-to-shape resizing.
    """

    template_r_pr = _find_template_run_properties(text_frame)
    paragraphs = list(text_frame.paragraphs)
    first_paragraph = paragraphs[0]
    txbody = text_frame._txBody  # noqa: SLF001

    for paragraph in paragraphs[1:]:
        txbody.remove(paragraph._p)  # noqa: SLF001

    _set_paragraph_text(first_paragraph, new_text, force=True)

    if template_r_pr is not None:
        run = first_paragraph.runs[0] if first_paragraph.runs else first_paragraph.add_run()
        existing_r_pr = run._r.find(qn("a:rPr"))  # noqa: SLF001
        if existing_r_pr is not None:
            run._r.remove(existing_r_pr)  # noqa: SLF001
        run._r.insert(0, template_r_pr)  # noqa: SLF001

    _apply_text_frame_fit(text_frame, new_text)


class PresentationTranslator:
    """Orchestrates the full translation of a PowerPoint presentation."""

    def __init__(
        self,
        translator: BaseTranslator,
        fallback_translator: BaseTranslator | None = None,
        exception_rules: list[ExceptionRule] | None = None,
        sample_chars_for_detection: int = 3000,
        remove_audio: bool = False,
    ):
        self.translator = translator
        self.fallback_translator = fallback_translator
        self.exception_rules = exception_rules or []
        self._pre_translation_rules, self._protected_rules = split_rules_by_mode(
            self.exception_rules
        )
        self.sample_chars_for_detection = sample_chars_for_detection
        self.remove_audio = remove_audio

    # -- Language detection ---------------------------------------------
    def _collect_detection_sample(self, prs) -> str:
        chunks: list[str] = []
        total = 0
        for slide in prs.slides:
            for shape, _in_group in _iter_shapes_recursive(slide.shapes, False):
                text = ""
                if getattr(shape, "has_text_frame", False):
                    text = shape.text_frame.text
                elif getattr(shape, "has_table", False):
                    text = " ".join(
                        cell.text_frame.text
                        for row in shape.table.rows
                        for cell in row.cells
                    )
                text = text.strip()
                if text:
                    chunks.append(text)
                    total += len(text)
                if total >= self.sample_chars_for_detection:
                    return "\n".join(chunks)
        return "\n".join(chunks)

    def _collect_first_slide_context(self, prs) -> str | None:
        if not prs.slides:
            return None

        chunks: list[str] = []
        first_slide = prs.slides[0]
        for shape, _in_group in _iter_shapes_recursive(first_slide.shapes, False):
            text = ""
            if getattr(shape, "has_text_frame", False):
                text = shape.text_frame.text
            elif getattr(shape, "has_table", False):
                text = " ".join(
                    cell.text_frame.text
                    for row in shape.table.rows
                    for cell in row.cells
                )
            text = text.strip()
            if text:
                chunks.append(text)

        if not chunks:
            return None

        context = "\n".join(chunks)
        if len(context) > 1500:
            context = context[:1500].rsplit(" ", 1)[0]
        return context.strip() or None

    # -- Collecting translation units ------------------------------------
    def _collect_jobs(self, prs, stats: TranslationStats) -> list[TranslationJob]:
        jobs: list[TranslationJob] = []
        slide_width = prs.slide_width
        slide_height = prs.slide_height

        for slide_index, slide in enumerate(prs.slides):
            for shape, in_group in _iter_shapes_recursive(slide.shapes, False):
                if getattr(shape, "has_table", False):
                    is_figure = in_group
                    for row in shape.table.rows:
                        for cell in row.cells:
                            jobs.extend(
                                self._collect_text_frame_jobs(
                                    cell.text_frame, is_figure, stats, slide_index=slide_index
                                )
                            )
                    continue

                if not getattr(shape, "has_text_frame", False):
                    continue

                text_frame = shape.text_frame
                if not text_frame.text.strip():
                    continue

                if in_group:
                    is_figure = True
                else:
                    width = None
                    height = None
                    try:
                        width = Length(shape.width) if shape.width else None
                        height = Length(shape.height) if shape.height else None
                    except Exception:  # pragma: no cover - shapes without their own geometry
                        pass
                    result = looks_like_figure_label(
                        text_frame.text, width, height, slide_width, slide_height
                    )
                    is_figure = result.is_figure
                    if is_figure:
                        stats.reasons_log.append(
                            f"'{text_frame.text[:40]!r}...' -> figure ({'; '.join(result.reasons)})"
                        )

                jobs.extend(self._collect_text_frame_jobs(text_frame, is_figure, stats, slide_index=slide_index))

            # Speaker notes are treated as regular content (never as a
            # figure): each paragraph is translated independently, keeping
            # its own run formatting (bold, italics, etc.), same as normal
            # slide text.
            if slide.has_notes_slide:
                notes_text_frame = slide.notes_slide.notes_text_frame
                if notes_text_frame.text.strip():
                    jobs.extend(
                        self._collect_text_frame_jobs(notes_text_frame, False, stats, slide_index=slide_index)
                    )

        return jobs

    def _collect_text_frame_jobs(
        self, text_frame, is_figure: bool, stats: TranslationStats, slide_index: int
    ) -> list[TranslationJob]:
        jobs: list[TranslationJob] = []
        if is_figure:
            pieces = [
                _paragraph_plain_text(p)
                for p in text_frame.paragraphs
                if p.text.strip() and _paragraph_has_text_run(p)
            ]
            joined = _clean_joined_text(" ".join(pieces))
            if joined:
                stats.figures_detected += 1
                jobs.append(
                    TranslationJob(
                        "figure",
                        text_frame,
                        joined,
                        text_frame=text_frame,
                        slide_index=slide_index,
                    )
                )
        else:
            for paragraph in text_frame.paragraphs:
                if not _paragraph_has_text_run(paragraph):
                    continue

                runs = []
                for run in paragraph.runs:
                    if not run.text.strip():
                        continue
                    raw_text = run.text
                    plain = _clean_joined_text(raw_text)
                    leading_ws = raw_text[: len(raw_text) - len(raw_text.lstrip())]
                    trailing_ws = raw_text[len(raw_text.rstrip()) :]
                    runs.append((run, plain, leading_ws, trailing_ws))
                if len(runs) > 1:
                    for run, plain, leading_ws, trailing_ws in runs:
                        job = TranslationJob(
                            "paragraph_run",
                            paragraph,
                            plain,
                            text_frame=text_frame,
                            slide_index=slide_index,
                        )
                        job.run_segments = [(run, plain, leading_ws, trailing_ws)]
                        jobs.append(job)
                    continue

                plain = _paragraph_plain_text(paragraph)
                if plain:
                    jobs.append(
                        TranslationJob(
                            "paragraph",
                            paragraph,
                            plain,
                            text_frame=text_frame,
                            slide_index=slide_index,
                        )
                    )
        return jobs

    # -- Translation and writing ------------------------------------------
    def translate(
        self,
        input_path: str,
        output_path: str,
        target_lang: str,
        source_lang: str | None = None,
    ) -> TranslationStats:
        stats = TranslationStats()
        prs = Presentation(input_path)
        stats.slides = len(prs.slides)

        # 1. Remove audio from all slides only when explicitly enabled.
        if self.remove_audio:
            for slide in prs.slides:
                stats.audio_removed += remove_audio_from_slide(slide)

        # 2. Auto-detect the source language if not specified.
        if source_lang:
            resolved_source_lang = source_lang
        else:
            sample = self._collect_detection_sample(prs)
            resolved_source_lang = detect_language(sample)
        stats.detected_source_lang = resolved_source_lang

        # 3. Collect all text units to translate.
        jobs = self._collect_jobs(prs, stats)

        # 4. Apply user-defined translation exceptions:
        #    - '<' (pre-translation) rules are substituted directly, since
        #      their destination text is expected to already be in the
        #      target language and can be translated along with the rest.
        #    - '!'/'~' (strict) rules are protected with placeholder tokens
        #      so the translator cannot alter them, then restored verbatim
        #      once translation is complete.
        #    Texts are grouped per slide so the translation unit is the slide.
        next_placeholder_index = 0
        for job in jobs:
            if self._pre_translation_rules:
                job.original_text = apply_exception_rules(
                    job.original_text, self._pre_translation_rules
                )
            if self._protected_rules:
                job.original_text, job.exception_replacements = mask_exception_rules(
                    job.original_text,
                    self._protected_rules,
                    start_index=next_placeholder_index,
                )
                next_placeholder_index += len(job.exception_replacements)

        unique_texts = {job.original_text for job in jobs}
        stats.unique_texts = len(unique_texts)
        context = self._collect_first_slide_context(prs)
        protected_replacements = {
            token: value
            for job in jobs
            for token, value in job.exception_replacements.items()
        }
        for provider in (self.translator, self.fallback_translator):
            if provider is not None:
                provider._protected_replacements = protected_replacements
        try:
            self.translator.on_presentation_start(context)
            translations: dict[str, str] = {}
            failed_texts = []
            slide_jobs: dict[int, list[TranslationJob]] = {}
            for job in jobs:
                slide_jobs.setdefault(job.slide_index, []).append(job)

            slide_count = len(prs.slides)
            for slide_index in range(slide_count):
                slide_job_list = slide_jobs.get(slide_index, [])
                if not slide_job_list:
                    continue

                slide_texts = list(dict.fromkeys(job.original_text for job in slide_job_list))
                logger.info(
                    "Translating slide %d/%d (%d text units)...",
                    slide_index + 1,
                    slide_count,
                    len(slide_texts),
                )
                slide_translations, slide_failed = self.translator.translate_slide(
                    slide_texts,
                    resolved_source_lang,
                    target_lang,
                    fallback=self.fallback_translator,
                    context=context,
                )
                translations.update(slide_translations)
                failed_texts.extend(slide_failed)
        finally:
            for provider in (self.translator, self.fallback_translator):
                if provider is not None:
                    provider._protected_replacements = {}
        stats.failed_texts = len(failed_texts)

        # 5. Write the translations back into the presentation.
        for job in jobs:
            translated_text = translations.get(job.original_text, job.original_text)
            translated_text = restore_exception_rules(
                translated_text, job.exception_replacements
            )
            if job.kind == "figure":
                _set_text_frame_single_text(job.target, translated_text)
                _apply_text_frame_fit(job.text_frame or job.target, translated_text)
            elif job.kind == "paragraph_run":
                original_run, _, leading_ws, trailing_ws = job.run_segments[0]
                original_run.text = f"{leading_ws}{translated_text}{trailing_ws}"
                _apply_text_frame_fit(job.text_frame, translated_text)
            else:
                _set_paragraph_text(job.target, translated_text)
                _apply_text_frame_fit(job.text_frame, translated_text)
            stats.translated_units += 1

        prs.save(output_path)
        return stats
