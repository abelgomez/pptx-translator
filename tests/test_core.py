import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import tempfile
import unittest
from unittest.mock import patch

from pptx_translator.exceptions_list import (
    ExceptionMode,
    apply_exception_rules,
    load_exception_rules,
    mask_exception_rules,
    restore_exception_rules,
    split_rules_by_mode,
)
from pptx import Presentation
from pptx_translator.cli import _resolve_log_level
from pptx_translator.figure_heuristics import looks_like_figure_label
from pptx_translator import media as media_module
from pptx_translator.pptx_processor import PresentationTranslator
from pptx_translator.translators.base import BaseTranslator, TranslationError
from pptx_translator.translators.factory import create_translator
from pptx_translator.config import Settings
from pptx.oxml.ns import qn
from pptx.enum.text import MSO_AUTO_SIZE
from lxml import etree
import logging


class ExceptionRulesTests(unittest.TestCase):
    def _rules_from_text(self, text):
        import tempfile
        from pathlib import Path

        with tempfile.TemporaryDirectory() as tmp_dir:
            path = Path(tmp_dir) / "exceptions.txt"
            path.write_text(text, encoding="utf-8")
            return load_exception_rules(path)

    def test_numeric_wildcard_is_replaced(self):
        rules = self._rules_from_text("~Práctica {num} = Practical Lesson {num}\n")
        _, protected = split_rules_by_mode(rules)
        masked, replacements = mask_exception_rules("Práctica 1.", protected)
        self.assertEqual(restore_exception_rules(masked, replacements), "Practical Lesson 1.")
        masked, replacements = mask_exception_rules("práctica 3", protected)
        self.assertEqual(restore_exception_rules(masked, replacements), "Practical Lesson 3")

    def test_default_mode_is_case_insensitive_protect(self):
        rules = self._rules_from_text("SGBD = RDBM\n")
        self.assertEqual(rules[0].mode, ExceptionMode.CASE_INSENSITIVE)

    def test_strict_mode_is_case_sensitive(self):
        rules = self._rules_from_text("!SGA = SGA\n")
        self.assertEqual(rules[0].mode, ExceptionMode.STRICT)
        self.assertIsNone(rules[0].pattern.match("sga"))
        self.assertIsNotNone(rules[0].pattern.match("SGA"))

    def test_pre_translation_mode_is_plain_substitution(self):
        rules = self._rules_from_text("<Práctica {num} = Practical Lesson {num}\n")
        pre_rules, protected_rules = split_rules_by_mode(rules)
        self.assertEqual(len(pre_rules), 1)
        self.assertEqual(len(protected_rules), 0)
        self.assertEqual(
            apply_exception_rules("Práctica 1.", pre_rules), "Practical Lesson 1."
        )

    def test_unrelated_text_untouched(self):
        rules = self._rules_from_text("<Práctica {num} = Practical Lesson {num}\n")
        text = "El SGBD Oracle es un sistema de gestión de bases de datos."
        self.assertEqual(apply_exception_rules(text, rules), text)

    def test_blank_and_comment_lines_are_ignored(self):
        rules = self._rules_from_text(
            "# a comment\n\n   \n<Práctica {num} = Practical Lesson {num}\n"
        )
        self.assertEqual(len(rules), 1)

    def test_mask_and_restore_preserves_exact_acronym(self):
        rules = self._rules_from_text("!SGA = SGA\n")
        masked, replacements = mask_exception_rules("SGA: System Global Area", rules)
        self.assertNotIn("SGA", masked)
        self.assertIn("SGA", replacements.values())
        self.assertEqual(
            restore_exception_rules(masked, replacements),
            "SGA: System Global Area",
        )

    def test_placeholder_tokens_are_purely_alphabetic(self):
        # The generated placeholders must contain no digits, underscores or
        # other punctuation, since those are what caused a previous scheme
        # to be mangled by the translation engine's detokenizer.
        rules = self._rules_from_text("!SGA = SGA\n!SGBD = SGBD\n!RDBM = RDBM\n")
        masked, replacements = mask_exception_rules(
            "SGA y SGBD y RDBM son siglas.", rules
        )
        for token in replacements:
            self.assertRegex(token, r"^[a-z]+$")

    def test_placeholder_tokens_are_unique_across_different_texts(self):
        rules = self._rules_from_text("~Práctica {num} = Practical Lesson {num}\n!SGA = SGA\n")

        masked_a, replacements_a = mask_exception_rules("Práctica 1", rules, start_index=0)
        masked_b, replacements_b = mask_exception_rules("SGA y práctica 2", rules, start_index=len(replacements_a))

        self.assertNotEqual(set(replacements_a), set(replacements_b))
        self.assertEqual(
            restore_exception_rules(masked_a, replacements_a),
            "Practical Lesson 1",
        )
        self.assertEqual(
            restore_exception_rules(masked_b, replacements_b),
            "SGA y Practical Lesson 2",
        )

    def test_restore_is_case_insensitive_for_mangled_casing(self):
        # Simulates a translator capitalizing the placeholder token (e.g.
        # because it starts a sentence), which must still be restored.
        rules = self._rules_from_text("!SGA = SGA\n")
        masked, replacements = mask_exception_rules("SGA: System Global Area", rules)
        token = next(iter(replacements))
        capitalized = masked.replace(token, token.capitalize())
        self.assertEqual(
            restore_exception_rules(capitalized, replacements),
            "SGA: System Global Area",
        )

    def test_invalid_line_raises(self):
        with self.assertRaises(ValueError):
            self._rules_from_text("this line has no separator\n")


class CLIInterruptTests(unittest.TestCase):
    def test_main_reports_abort_without_stacktrace(self):
        import pptx_translator.cli as cli_module

        with patch("pptx_translator.cli.build_arg_parser", side_effect=KeyboardInterrupt):
            with patch("sys.stderr", new_callable=io.StringIO) as stderr:
                exit_code = cli_module.main([])

        self.assertEqual(exit_code, 130)
        self.assertIn("Execution aborted by the user", stderr.getvalue())


class LogLevelResolutionTests(unittest.TestCase):
    def test_no_flag_uses_configured_level(self):
        self.assertEqual(_resolve_log_level(0, "DEBUG"), logging.DEBUG)
        self.assertEqual(_resolve_log_level(0, "INFO"), logging.INFO)
        self.assertEqual(_resolve_log_level(0, "ERROR"), logging.ERROR)
        self.assertEqual(_resolve_log_level(0, "WARNING"), logging.WARNING)

    def test_no_flag_and_no_config_defaults_to_warning(self):
        self.assertEqual(_resolve_log_level(0, None), logging.WARNING)

    def test_invalid_configured_level_falls_back_to_warning(self):
        self.assertEqual(_resolve_log_level(0, "NOT_A_LEVEL"), logging.WARNING)

    def test_verbose_flags_override_configured_level(self):
        self.assertEqual(_resolve_log_level(1, "ERROR"), logging.INFO)
        self.assertEqual(_resolve_log_level(2, "ERROR"), logging.DEBUG)


class FigureHeuristicTests(unittest.TestCase):
    def test_short_multiline_small_box_is_figure(self):
        result = looks_like_figure_label(
            "Ordenador con\nsoftware Oracle",
            shape_width_emu=900000,
            shape_height_emu=500000,
            slide_width_emu=12192000,
            slide_height_emu=6858000,
        )
        self.assertTrue(result.is_figure)

    def test_long_sentence_is_not_figure(self):
        result = looks_like_figure_label(
            "Instancia Oracle: SGA (System Global Area) + Procesos de fondo. "
            "Es el medio para acceder a los datos de la base de datos.",
            shape_width_emu=9000000,
            shape_height_emu=3000000,
            slide_width_emu=12192000,
            slide_height_emu=6858000,
        )
        self.assertFalse(result.is_figure)

    def test_wide_box_with_few_words_keeps_line_breaks(self):
        result = looks_like_figure_label(
            "Primera linea\nSegunda linea",
            shape_width_emu=7000000,
            shape_height_emu=500000,
            slide_width_emu=12192000,
            slide_height_emu=6858000,
        )
        self.assertFalse(result.is_figure)

    def test_single_line_is_not_figure(self):
        result = looks_like_figure_label(
            "Objetivos:",
            shape_width_emu=900000,
            shape_height_emu=300000,
            slide_width_emu=12192000,
            slide_height_emu=6858000,
        )
        self.assertFalse(result.is_figure)


class ParagraphRunFormattingTests(unittest.TestCase):
    def test_styled_note_runs_keep_their_own_formatting(self):
        class _StubTranslator(BaseTranslator):
            name = "stub"

            def _translate_slide(
                self,
                texts,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                return {text: f"TR:{text}" for text in texts}

        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "notes_test.pptx"
            output_path = Path(tmp_dir) / "notes_test_en.pptx"

            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            notes = slide.notes_slide.notes_text_frame
            paragraph = notes.paragraphs[0]

            first_run = paragraph.add_run()
            first_run.text = "Concepto"
            first_run.font.bold = True

            second_run = paragraph.add_run()
            second_run.text = " importante"
            second_run.font.italic = True

            prs.save(input_path)

            translator = _StubTranslator()
            PresentationTranslator(translator).translate(
                str(input_path), str(output_path), target_lang="en", source_lang="es"
            )

            output_prs = Presentation(output_path)
            output_slide = output_prs.slides[0]
            paragraph = output_slide.notes_slide.notes_text_frame.paragraphs[0]
            self.assertEqual(paragraph.runs[0].text, "TR:Concepto")
            self.assertTrue(paragraph.runs[0].font.bold)
            self.assertEqual(paragraph.runs[1].text, " TR:importante")
            self.assertTrue(paragraph.runs[1].font.italic)

    def test_regular_text_boxes_enable_word_wrap_and_autosize(self):
        class _StubTranslator(BaseTranslator):
            name = "stub"

            def _translate_slide(
                self,
                texts,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                return {text: f"TR:{text}" for text in texts}

        with tempfile.TemporaryDirectory() as tmp_dir:
            input_path = Path(tmp_dir) / "textbox_test.pptx"
            output_path = Path(tmp_dir) / "textbox_test_en.pptx"

            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            shape = slide.shapes.add_textbox(10, 10, 200, 100)
            shape.text_frame.text = "Concepto importante para la prueba"

            prs.save(input_path)

            PresentationTranslator(_StubTranslator()).translate(
                str(input_path), str(output_path), target_lang="en", source_lang="es"
            )

            output_prs = Presentation(output_path)
            output_shape = output_prs.slides[0].shapes[0]
            self.assertTrue(output_shape.text_frame.word_wrap)
            self.assertEqual(
                output_shape.text_frame.auto_size,
                MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
            )

    def test_first_slide_context_is_logged_once_when_a_presentation_starts(self):
        from pptx_translator.translators.openai import OpenAITranslator

        translator = OpenAITranslator(api_key="demo", api_base_url="https://example.test/v1")
        context = "Título del estudio\n\nPrimera línea\nSegunda línea"
        formatted_context = "Título del estudio / Primera línea / Segunda línea"

        with patch("pptx_translator.translators.openai.logger") as mock_logger:
            translator.on_presentation_start(context)
            translator.on_presentation_start(context)

        self.assertEqual(mock_logger.info.call_count, 2)
        self.assertTrue(
            any(
                "Using first-slide presentation context for OpenAI-compatible translation" in str(call)
                for call in mock_logger.info.call_args_list
            )
        )
        self.assertTrue(
            any(
                formatted_context in str(call)
                for call in mock_logger.info.call_args_list
            )
        )
        self.assertTrue(
            all("\n" not in str(call) for call in mock_logger.info.call_args_list)
        )

    def test_openai_system_content_includes_protected_replacements(self):
        from pptx_translator.translators.openai import OpenAITranslator

        translator = OpenAITranslator(api_key="demo", api_base_url="https://example.test/v1")
        translator._protected_replacements = {
            "zqkpptxaxvxq": "SGA",
            "zqkpptxbyvxq": "SGA",
            "zqkpptxcvvxq": "Practical Lesson 1",
            "zqkpptxdxvxq": "Practical Lesson 1",
        }

        content = translator._build_system_content("Demo context")

        self.assertIn("zqkpptxaxvxq -> SGA", content)
        self.assertIn("zqkpptxcvvxq -> Practical Lesson 1", content)
        self.assertEqual(content.count(" -> SGA"), 1)
        self.assertEqual(content.count(" -> Practical Lesson 1"), 1)

    def test_openai_slide_translation_uses_json_payload_for_batch_requests(self):
        from pptx_translator.translators.openai import OpenAITranslator

        translator = OpenAITranslator(api_key="demo", api_base_url="https://example.test/v1")
        with patch("pptx_translator.translators.openai.requests.post") as mock_post:
            mock_post.return_value.status_code = 200
            mock_post.return_value.json.return_value = {
                "choices": [{"message": {"content": '{"translations":[{"id":"Hola","text":"Hello"}]}'}}]
            }

            result = translator._translate_slide(["Hola"], "es", "en")

        self.assertEqual(result, {"Hola": "Hello"})
        payload = mock_post.call_args.kwargs["json"]
        self.assertIn("single key 'translations'", payload["messages"][1]["content"])
        self.assertIn('"items":', payload["messages"][1]["content"])
        self.assertIn('"id": "0"', payload["messages"][1]["content"])

    def test_openai_slide_retries_with_single_text_calls_when_json_parse_fails(self):
        from pptx_translator.translators.openai import OpenAITranslator

        translator = OpenAITranslator(api_key="demo", api_base_url="https://example.test/v1")
        responses = [
            unittest.mock.Mock(status_code=200, json=unittest.mock.Mock(return_value={"choices": [{"message": {"content": "not-json"}}]})),
            unittest.mock.Mock(status_code=200, json=unittest.mock.Mock(return_value={"choices": [{"message": {"content": "Hello"}}]})),
        ]

        with patch("pptx_translator.translators.openai.requests.post", side_effect=responses) as mock_post:
            result = translator._translate_slide(["Hola"], "es", "en")

        self.assertEqual(result, {"Hola": "Hello"})
        self.assertEqual(mock_post.call_count, 2)
        self.assertIn("Translate this presentation text", mock_post.call_args_list[1].kwargs["json"]["messages"][1]["content"])

    def test_openai_single_text_request_avoids_json_formatting(self):
        from pptx_translator.translators.openai import OpenAITranslator

        translator = OpenAITranslator(api_key="demo", api_base_url="https://example.test/v1")
        with patch("pptx_translator.translators.openai.requests.post") as mock_post:
            mock_post.return_value.status_code = 200
            mock_post.return_value.json.return_value = {
                "choices": [{"message": {"content": "Hello"}}]
            }

            translator._translate_single_text("Práctica 1", "es", "en")

        user_content = mock_post.call_args.kwargs["json"]["messages"][1]["content"]
        self.assertNotIn("single key 'translations'", user_content)
        self.assertNotIn("```json", user_content.lower())
        self.assertIn("Return only the translated text", user_content)

    def test_local_translator_keeps_original_text_when_individual_item_fails(self):
        class _LocalLikeTranslator(BaseTranslator):
            name = "local"

            def _translate_slide(
                self,
                texts,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                raise TranslationError("unexpected batch call")

            def _translate_single_text(
                self,
                text,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                raise TranslationError("local item failed")

        class _FallbackTranslator(BaseTranslator):
            name = "fallback"

            def _translate_slide(
                self,
                texts,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                return {text: f"TR:{text}" for text in texts}

            def _translate_single_text(
                self,
                text,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                return f"TR:{text}"

        translator = _LocalLikeTranslator()
        results, failed = translator.translate_slide(["Práctica"], "es", "en", fallback=_FallbackTranslator())

        self.assertEqual(results, {"Práctica": "Práctica"})
        self.assertEqual(failed, ["Práctica"])

    def test_remote_translator_uses_local_fallback_after_single_item_failures(self):
        from pptx_translator.translators.base import TranslationError

        class _RemoteLikeTranslator(BaseTranslator):
            name = "openai"

            def _translate_slide(
                self,
                texts,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                raise TranslationError("batch failed")

            def _translate_single_text(
                self,
                text,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                raise TranslationError("single item failed")

        class _LocalFallback(BaseTranslator):
            name = "local"

            def _translate_slide(
                self,
                texts,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                return {text: text for text in texts}

            def _translate_single_text(
                self,
                text,
                source_lang,
                target_lang,
                context=None,
                exception_rules=None,
            ):
                return text

        translator = _RemoteLikeTranslator()
        results, failed = translator.translate_slide(["Práctica"], "es", "en", fallback=_LocalFallback())

        self.assertEqual(results, {"Práctica": "Práctica"})
        self.assertEqual(failed, [])


class CliLoggingTests(unittest.TestCase):
    def test_selected_provider_log_uses_effective_provider_only(self):
        from pptx_translator.cli import _log_active_configuration
        from pptx_translator.config import Settings

        logger = unittest.mock.Mock()
        logger.isEnabledFor.return_value = True
        _log_active_configuration(logger, Settings(provider="openai", remove_audio=False), "local")

        logger.info.assert_any_call("Selected translation provider: %s", "local")
        logger.info.assert_any_call("Audio removal enabled: %s", False)

    def test_argos_internal_logging_mutes_stanza_loggers(self):
        from pptx_translator.translators.local_argos import _sync_argos_internal_logging

        logging.basicConfig(level=logging.WARNING, force=True)
        stanza_logger = logging.getLogger("stanza")
        stanza_logger.disabled = False
        stanza_logger.setLevel(logging.INFO)

        _sync_argos_internal_logging()

        self.assertTrue(stanza_logger.disabled)
        self.assertEqual(logging.getLogger("argostranslate").getEffectiveLevel(), logging.CRITICAL)
        self.assertEqual(logging.getLogger("argostranslate.utils").getEffectiveLevel(), logging.CRITICAL)


class CliBatchProcessingTests(unittest.TestCase):
    def test_should_skip_translation_when_file_already_matches_target_suffix(self):
        from pptx_translator.cli import _should_skip_translation

        self.assertTrue(_should_skip_translation(Path("demo_en.pptx"), "en"))
        self.assertTrue(_should_skip_translation(Path("demo_EN.pptx"), "en"))
        self.assertFalse(_should_skip_translation(Path("demo.pptx"), "en"))
        self.assertFalse(_should_skip_translation(Path("demo_es.pptx"), "en"))

    def test_skips_files_already_named_for_target_language_and_logs_warning(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            already_translated = root / "presentation_en.pptx"
            already_translated.write_bytes(b"pptx")

            with patch("pptx_translator.cli.create_translator") as mock_factory, patch(
                "pptx_translator.cli.PresentationTranslator"
            ) as mock_translator_cls, patch("pptx_translator.cli.logging.getLogger") as mock_get_logger:
                mock_logger = unittest.mock.Mock()
                mock_get_logger.return_value = mock_logger
                mock_factory.return_value.name = "stub"

                result = __import__("pptx_translator.cli", fromlist=["main"]).main(
                    [str(root), "-t", "en"]
                )

            self.assertEqual(result, 0)
            self.assertEqual(mock_translator_cls.call_count, 0)
            mock_logger.warning.assert_any_call(
                "Skipping file '%s': it already ends with the target-language suffix '_%s'.",
                already_translated,
                "en",
            )

    def test_dry_run_does_not_write_output_or_call_translator(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            source = root / "presentation.pptx"
            target = root / "presentation_en.pptx"
            source.write_bytes(b"pptx")
            target.write_bytes(b"existing")

            with patch("pptx_translator.cli.create_translator") as mock_factory, patch(
                "pptx_translator.cli.PresentationTranslator"
            ) as mock_translator_cls:
                mock_factory.return_value.name = "stub"

                result = __import__("pptx_translator.cli", fromlist=["main"]).main(
                    [str(source), "-t", "en", "--dry-run"]
                )

            self.assertEqual(result, 0)
            self.assertEqual(mock_translator_cls.call_count, 0)
            self.assertEqual(target.read_bytes(), b"existing")

    def test_directory_without_recursive_only_processes_top_level_ppts(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            top = root / "top.pptx"
            nested_dir = root / "nested"
            nested_dir.mkdir()
            nested = nested_dir / "inner.pptx"
            top.write_bytes(b"pptx")
            nested.write_bytes(b"pptx")
            (root / "other.txt").write_text("ignore me", encoding="utf-8")

            with patch("pptx_translator.cli.create_translator") as mock_factory, patch(
                "pptx_translator.cli.PresentationTranslator"
            ) as mock_translator_cls:
                mock_factory.return_value.name = "stub"
                mock_translator_cls.return_value.translate.return_value = type(
                    "Stats",
                    (),
                    {
                        "detected_source_lang": "es",
                        "slides": 1,
                        "audio_removed": 0,
                        "figures_detected": 0,
                        "unique_texts": 1,
                        "translated_units": 1,
                        "failed_texts": 0,
                        "reasons_log": [],
                    },
                )()

                result = __import__("pptx_translator.cli", fromlist=["main"]).main(
                    [str(root), "-t", "en"]
                )

            self.assertEqual(result, 0)
            self.assertEqual(mock_translator_cls.call_count, 1)
            self.assertEqual(str(mock_translator_cls.call_args[0][0].name), "stub")

    def test_directory_with_recursive_processes_nested_ppts(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            nested_dir = root / "nested"
            nested_dir.mkdir()
            (root / "top.pptx").write_bytes(b"pptx")
            (nested_dir / "inner.pptx").write_bytes(b"pptx")

            with patch("pptx_translator.cli.create_translator") as mock_factory, patch(
                "pptx_translator.cli.PresentationTranslator"
            ) as mock_translator_cls:
                mock_factory.return_value.name = "stub"
                mock_translator_cls.return_value.translate.return_value = type(
                    "Stats",
                    (),
                    {
                        "detected_source_lang": "es",
                        "slides": 1,
                        "audio_removed": 0,
                        "figures_detected": 0,
                        "unique_texts": 1,
                        "translated_units": 1,
                        "failed_texts": 0,
                        "reasons_log": [],
                    },
                )()

                result = __import__("pptx_translator.cli", fromlist=["main"]).main(
                    [str(root), "-t", "en", "-r"]
                )

            self.assertEqual(result, 0)
            self.assertEqual(mock_translator_cls.call_count, 2)


class _FakeSlide:
    """Minimal stand-in for a python-pptx Slide, exposing only ``_element``."""

    def __init__(self, element):
        self._element = element


class TranslatorFactoryTests(unittest.TestCase):
    def test_factory_uses_local_provider_without_exposed_configuration(self):
        translator = create_translator(Settings(), provider="local")
        self.assertEqual(translator.name, "local")

    def test_factory_supports_openai_compatible_provider(self):
        provider = create_translator(Settings(provider="openai", api_key="demo"), provider="openai")
        self.assertEqual(provider.name, "openai")


class OrphanedTimingCleanupTests(unittest.TestCase):
    """Covers the fix for PPTX files PowerPoint reports as needing repair
    after audio shapes were removed: leftover <p:timing> entries still
    referencing the deleted shape's id must be pruned too."""

    def _build_slide_with_timing(self, referenced_spid: str, other_spid: str | None = None):
        p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
        sld = etree.SubElement(
            etree.Element(f"{{{p_ns}}}root"), f"{{{p_ns}}}sld"
        )
        timing = etree.SubElement(sld, f"{{{p_ns}}}timing")
        tn_lst = etree.SubElement(timing, f"{{{p_ns}}}tnLst")

        def _add_par_for(spid):
            par = etree.SubElement(tn_lst, f"{{{p_ns}}}par")
            c_tn = etree.SubElement(par, f"{{{p_ns}}}cTn")
            child_tn_lst = etree.SubElement(c_tn, f"{{{p_ns}}}childTnLst")
            cond = etree.SubElement(child_tn_lst, f"{{{p_ns}}}cond")
            tgt_el = etree.SubElement(cond, f"{{{p_ns}}}tgtEl")
            etree.SubElement(tgt_el, f"{{{p_ns}}}spTgt", spid=spid)

        _add_par_for(referenced_spid)
        if other_spid is not None:
            _add_par_for(other_spid)

        return _FakeSlide(sld), timing

    def test_dangling_timing_entry_is_removed(self):
        slide, timing = self._build_slide_with_timing("5")
        media_module._remove_orphaned_timing_nodes(slide, ["5"])
        # No dangling spTgt should remain, and since it was the only
        # animation, the whole (now useless) <p:timing> tree is dropped.
        self.assertIsNone(slide._element.find(qn("p:timing")))

    def test_unrelated_timing_entries_are_kept(self):
        slide, timing = self._build_slide_with_timing("5", other_spid="7")
        media_module._remove_orphaned_timing_nodes(slide, ["5"])
        remaining = timing.findall(f".//{qn('p:spTgt')}")
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0].attrib.get("spid"), "7")

    def test_no_removed_ids_is_a_no_op(self):
        slide, timing = self._build_slide_with_timing("5")
        media_module._remove_orphaned_timing_nodes(slide, [])
        self.assertIsNotNone(slide._element.find(qn("p:timing")))


if __name__ == "__main__":
    unittest.main()
