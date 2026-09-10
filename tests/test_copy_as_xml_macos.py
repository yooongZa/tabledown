"""macOS direct Excel-selection table-to-XML action regression tests."""

from __future__ import annotations

import threading
import unittest
import xml.etree.ElementTree as ET
from types import SimpleNamespace
from unittest.mock import Mock, patch

from tablemark.app import TabledownApp
from tablemark.clipboard import ClipboardChangedError, ClipboardWriteError
from tablemark.converter.table_xml import TableXmlTooLargeError
from tablemark.excel_formula import (
    DISPLAY_OVERFLOW,
    EXCEL_NOT_RUNNING,
    MULTIPLE_AREAS,
    PARTIAL_MERGE,
    ExcelFormulaError,
)
from tablemark.excel_table import ExcelTableSelection
from tablemark.i18n import t


def selection(values=("상품", "값", "사과", "2")):
    return ExcelTableSelection(
        workbook="Book.xlsx",
        sheet="Sheet1",
        address="$A$1:$B$2",
        row_count=2,
        column_count=2,
        values=tuple(values),
        merge_areas=(),
    )


class CopyAsXmlActionTests(unittest.TestCase):
    @staticmethod
    def _app():
        app = SimpleNamespace(
            fill_blanks=False,
            lang="ko",
            _clipboard_operation_lock=threading.Lock(),
            _explicit_export_lock=threading.Lock(),
            _explicit_export_active=None,
            _stop_watcher=threading.Event(),
            _clipboard_table_model=TabledownApp._clipboard_table_model,
            _flash_icon_success=Mock(),
            _safe_alert=Mock(),
            copy_xml_item=SimpleNamespace(title="", _menuitem=Mock()),
            copy_excel_formulas_item=SimpleNamespace(title="", _menuitem=Mock()),
            copy_markdown_item=SimpleNamespace(title="", _menuitem=Mock()),
        )
        for name in (
            "copy_as_xml",
            "copy_as_markdown",
            "_start_explicit_export",
            "_run_explicit_export",
            "_perform_explicit_export",
            "_finish_explicit_export",
            "_alert_explicit_export_error",
            "_set_explicit_export_busy",
            "_update_explicit_export_menu",
            "_augment_clipboard",
        ):
            setattr(app, name, getattr(TabledownApp, name).__get__(app))
        return app

    @staticmethod
    def _run(app):
        with (
            patch("tablemark.app.clipboard_change_count", return_value=17),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=lambda callback, *args: callback(*args),
            ),
        ):
            app.copy_as_xml(None)

    def test_success_reads_excel_selection_without_reading_clipboard(self):
        app = self._app()
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=selection(),
            ) as reader,
            patch("tablemark.app.read_clipboard") as clipboard_reader,
            patch("tablemark.app.write_text_only_clipboard") as text_only_writer,
            patch("tablemark.app.write_clipboard") as preserving_writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        xml = text_only_writer.call_args.args[0]
        self.assertTrue(xml.startswith("<표 "))
        self.assertIn('형식버전="2"', xml)
        self.assertIn('통합문서="Book.xlsx"', xml)
        self.assertIn('주소="$A$1:$B$2"', xml)
        self.assertIn('병합범위=""', xml)
        self.assertIn('<열 n="상품">사과</열>', xml)
        reader.assert_called_once_with()
        clipboard_reader.assert_not_called()
        text_only_writer.assert_called_once_with(
            xml,
            mark_generated=True,
            expected_change_count=17,
        )
        preserving_writer.assert_not_called()
        app._safe_alert.assert_not_called()
        app._flash_icon_success.assert_called_once_with()
        self.assertIsNone(app.title)
        logger.assert_called_once_with("copied selected Excel table as XML")

    def test_header_only_selection_leaves_clipboard_untouched(self):
        app = self._app()
        header_only = ExcelTableSelection(
            workbook="Book.xlsx",
            sheet="Sheet1",
            address="$A$1:$B$1",
            row_count=1,
            column_count=2,
            values=("상품", "값"),
            merge_areas=(),
        )
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=header_only,
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: no_table"
        )
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.no_table", "ko"),
        )

    def test_selection_errors_leave_clipboard_untouched_and_are_localized(self):
        for code in (
            EXCEL_NOT_RUNNING,
            MULTIPLE_AREAS,
            PARTIAL_MERGE,
            DISPLAY_OVERFLOW,
        ):
            with self.subTest(code=code):
                app = self._app()
                with (
                    patch(
                        "tablemark.app.read_stable_selected_excel_table",
                        side_effect=ExcelFormulaError(code),
                    ),
                    patch("tablemark.app.write_text_only_clipboard") as writer,
                    patch("tablemark.app.log") as logger,
                ):
                    self._run(app)

                writer.assert_not_called()
                app._flash_icon_success.assert_not_called()
                logger.assert_called_once_with(
                    f"copy selected Excel table as XML failed: {code}"
                )
                app._safe_alert.assert_called_once_with(
                    "Tabledown", t(f"table.error.{code}", "ko")
                )

    def test_writer_failure_does_not_report_success_or_log_private_content(self):
        app = self._app()
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=selection(),
            ),
            patch(
                "tablemark.app.write_text_only_clipboard",
                side_effect=OSError("PRIVATE_CLIPBOARD_DETAIL"),
            ),
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: OSError"
        )
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.execution_failed", "ko"),
        )

    def test_verified_clipboard_write_failure_has_actionable_error(self):
        app = self._app()
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=selection(),
            ),
            patch(
                "tablemark.app.write_text_only_clipboard",
                side_effect=ClipboardWriteError("clipboard_write_failed"),
            ),
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: clipboard_write_failed"
        )
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.clipboard_write_failed", "ko"),
        )

    def test_output_limit_fails_before_clipboard_write(self):
        app = self._app()
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=selection(),
            ),
            patch(
                "tablemark.app.model_to_xml",
                side_effect=TableXmlTooLargeError("too large"),
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: output_too_large"
        )
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.output_too_large", "ko"),
        )

    def test_xml_forbidden_cell_character_fails_before_clipboard_write(self):
        app = self._app()
        invalid = selection(("상품", "값", "사과", "bad\x01value"))
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=invalid,
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: ValueError"
        )
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.execution_failed", "ko"),
        )

    def test_menu_help_and_all_selection_errors_are_translated(self):
        self.assertEqual(
            t("menu.copy_xml", "ko"), "XML 변환 복사"
        )
        self.assertEqual(
            t("menu.copy_xml", "en"),
            "Copy as XML",
        )
        for language in ("ko", "en"):
            help_text = t("help.message", language)
            self.assertIn(t("menu.copy_xml", language), help_text)
            self.assertIn("⌘⌃X", help_text)
            for code in (
                EXCEL_NOT_RUNNING,
                "no_selection",
                MULTIPLE_AREAS,
                "too_many_cells",
                "too_much_text",
                "selection_changed",
                "automation_denied",
                PARTIAL_MERGE,
                DISPLAY_OVERFLOW,
                "execution_failed",
                "invalid_response",
                "no_table",
                "clipboard_write_failed",
                "output_too_large",
            ):
                key = f"table.error.{code}"
                self.assertNotEqual(t(key, language), key)
            self.assertNotEqual(
                t("table.error.clipboard_changed", language),
                "table.error.clipboard_changed",
            )

    def test_active_explicit_export_defers_watcher_without_touching_clipboard(self):
        app = self._app()
        app._explicit_export_active = "table"
        with (
            patch("tablemark.app.read_clipboard") as reader,
            patch("tablemark.app.write_clipboard") as writer,
        ):
            processed = app._augment_clipboard()

        self.assertFalse(processed)
        reader.assert_not_called()
        writer.assert_not_called()

    def test_watcher_rechecks_export_state_inside_clipboard_lock(self):
        app = self._app()

        class ExportStartsOnEnter:
            def __enter__(inner_self):
                app._explicit_export_active = "formulas"

            def __exit__(inner_self, exc_type, exc, traceback):
                return False

        app._clipboard_operation_lock = ExportStartsOnEnter()
        with (
            patch("tablemark.app.read_clipboard") as reader,
            patch("tablemark.app.write_clipboard") as writer,
        ):
            processed = app._augment_clipboard()

        self.assertFalse(processed)
        reader.assert_not_called()
        writer.assert_not_called()

    def test_watcher_does_not_overwrite_a_newer_external_copy(self):
        app = self._app()
        app._converted_clipboard = Mock(
            return_value={
                "text": "| A |",
                "html": "<table><tr><td>A</td></tr></table>",
                "drop_types": set(),
            }
        )
        content = object()
        with (
            patch("tablemark.app.clipboard_change_count", return_value=41),
            patch("tablemark.app.read_clipboard", return_value=content) as reader,
            patch(
                "tablemark.app.write_clipboard",
                side_effect=ClipboardChangedError("clipboard_changed"),
            ) as writer,
            patch("tablemark.app.AppHelper.callAfter") as call_after,
            patch("tablemark.app.log") as logger,
        ):
            processed = app._augment_clipboard()

        self.assertFalse(processed)
        reader.assert_called_once_with()
        app._converted_clipboard.assert_called_once_with(content)
        writer.assert_called_once_with(
            text="| A |",
            html="<table><tr><td>A</td></tr></table>",
            drop_types=set(),
            mark_generated=True,
            expected_change_count=41,
        )
        call_after.assert_not_called()
        logger.assert_not_called()

    def test_action_returns_after_scheduling_main_thread_export(self):
        app = self._app()
        scheduled = []

        with (
            patch("tablemark.app.clipboard_change_count", return_value=4),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=lambda callback, *args: scheduled.append((callback, args)),
            ),
            patch("tablemark.app.read_stable_selected_excel_table") as reader,
        ):
            app.copy_as_xml(None)

        reader.assert_not_called()
        self.assertEqual(len(scheduled), 1)
        self.assertIs(scheduled[0][0].__self__, app)
        self.assertEqual(scheduled[0][0].__func__, TabledownApp._run_explicit_export)
        self.assertEqual(app.title, "…")
        self.assertEqual(app.copy_xml_item.title, t("menu.copy_xml_busy", "ko"))
        app.copy_xml_item._menuitem.setEnabled_.assert_called_with(False)
        app.copy_excel_formulas_item._menuitem.setEnabled_.assert_called_with(False)

    def test_main_thread_schedule_failure_restores_menu_and_releases_gate(self):
        app = self._app()

        with (
            patch("tablemark.app.clipboard_change_count", return_value=4),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=OSError("main-loop unavailable"),
            ),
            patch("tablemark.app.log") as logger,
        ):
            app.copy_as_xml(None)

        self.assertIsNone(app._explicit_export_active)
        self.assertIsNone(app.title)
        self.assertEqual(app.copy_xml_item.title, t("menu.copy_xml", "ko"))
        self.assertTrue(app._explicit_export_lock.acquire(blocking=False))
        app._explicit_export_lock.release()
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.execution_failed", "ko"),
        )
        logger.assert_called_once_with("start explicit XML export failed: OSError")

    def test_second_export_is_blocked_with_feedback(self):
        app = self._app()

        with (
            patch("tablemark.app.clipboard_change_count", return_value=4),
            patch("tablemark.app.AppHelper.callAfter"),
        ):
            app.copy_as_xml(None)
            app.copy_as_xml(None)

        app._safe_alert.assert_called_once_with(
            t("export.error_title", "ko"),
            t("export.error.in_progress", "ko"),
        )

    def test_newer_clipboard_cancels_late_write_and_success(self):
        app = self._app()
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=selection(),
            ),
            patch(
                "tablemark.app.write_text_only_clipboard",
                side_effect=ClipboardChangedError("clipboard_changed"),
            ) as writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(app)

        writer.assert_called_once()
        app._flash_icon_success.assert_not_called()
        app._safe_alert.assert_called_once_with(
            "Tabledown",
            t("table.error.clipboard_changed", "ko"),
        )
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: clipboard_changed"
        )

    def test_fill_blanks_value_is_snapshotted_when_action_starts(self):
        app = self._app()
        app.fill_blanks = True
        scheduled = []

        with (
            patch("tablemark.app.clipboard_change_count", return_value=4),
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=selection(),
            ),
            patch(
                "tablemark.app.forward_fill_key_columns",
                side_effect=lambda rows: rows,
            ) as fill,
            patch("tablemark.app.write_text_only_clipboard"),
            patch("tablemark.app.log"),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=lambda callback, *args: scheduled.append((callback, args)),
            ),
        ):
            app.copy_as_xml(None)
            app.fill_blanks = False
            callback, args = scheduled[0]
            callback(*args)

        fill.assert_called_once()
        self.assertTrue(app._explicit_export_lock.acquire(blocking=False))
        app._explicit_export_lock.release()

    def test_fill_blanks_records_exact_generated_source_cells(self):
        app = self._app()
        app.fill_blanks = True
        grouped = ExcelTableSelection(
            workbook="Book.xlsx",
            sheet="Sheet1",
            address="$A$1:$C$3",
            row_count=3,
            column_count=3,
            values=(
                "그룹",
                "항목",
                "값",
                "A",
                "x",
                "1",
                None,
                "y",
                "2",
            ),
            merge_areas=(),
        )
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=grouped,
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log"),
        ):
            self._run(app)

        root = ET.fromstring(writer.call_args.args[0])
        self.assertEqual(root.attrib["빈칸채움"], "적용")
        self.assertEqual(root.attrib["빈칸채움수"], "1")
        self.assertEqual(root.attrib["빈칸채움셀"], "$A$3")
        self.assertIn("A", [cell.text for cell in root.findall(".//열")])

    def test_fill_blanks_deduplicates_one_blank_merged_source(self):
        app = self._app()
        app.fill_blanks = True
        blank_merge = ExcelTableSelection(
            workbook="Book.xlsx",
            sheet="Sheet1",
            address="$A$1:$C$4",
            row_count=4,
            column_count=3,
            values=(
                "그룹",
                "항목",
                "값",
                "A",
                "x",
                "1",
                None,
                "y",
                "2",
                None,
                "z",
                "3",
            ),
            merge_areas=("$A$3:$A$4",),
        )
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_table",
                return_value=blank_merge,
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log"),
        ):
            self._run(app)

        root = ET.fromstring(writer.call_args.args[0])
        self.assertEqual(root.attrib["빈칸채움수"], "1")
        self.assertEqual(root.attrib["빈칸채움셀"], "$A$3")

    def test_quit_before_scheduled_export_cancels_without_excel_or_ui_feedback(self):
        app = self._app()
        app._stop_watcher.set()
        with (
            patch("tablemark.app.read_stable_selected_excel_table") as reader,
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log"),
        ):
            self._run(app)

        reader.assert_not_called()
        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()
        app._safe_alert.assert_not_called()
        self.assertTrue(app._explicit_export_lock.acquire(blocking=False))
        app._explicit_export_lock.release()


if __name__ == "__main__":
    unittest.main()
