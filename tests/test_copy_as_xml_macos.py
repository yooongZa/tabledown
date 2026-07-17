"""macOS direct Excel-selection table-to-XML action regression tests."""

from __future__ import annotations

import threading
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from tablemark.app import TabledownApp
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
        return SimpleNamespace(
            fill_blanks=False,
            lang="ko",
            _clipboard_operation_lock=threading.Lock(),
            _clipboard_table_model=TabledownApp._clipboard_table_model,
            _flash_icon_success=Mock(),
            _safe_alert=Mock(),
        )

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
            TabledownApp.copy_as_xml(app, None)

        xml = text_only_writer.call_args.args[0]
        self.assertTrue(xml.startswith("<표>"))
        self.assertIn('<열 n="상품">사과</열>', xml)
        reader.assert_called_once_with()
        clipboard_reader.assert_not_called()
        text_only_writer.assert_called_once_with(xml, mark_generated=True)
        preserving_writer.assert_not_called()
        app._safe_alert.assert_not_called()
        app._flash_icon_success.assert_called_once_with()
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
            TabledownApp.copy_as_xml(app, None)

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
                    TabledownApp.copy_as_xml(app, None)

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
            TabledownApp.copy_as_xml(app, None)

        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with(
            "copy selected Excel table as XML failed: OSError"
        )
        app._safe_alert.assert_called_once_with(
            t("table.error_title", "ko"),
            t("table.error.execution_failed", "ko"),
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
            TabledownApp.copy_as_xml(app, None)

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
        self.assertEqual(t("menu.copy_xml", "ko"), "선택한 표를 XML로 복사")
        self.assertEqual(t("menu.copy_xml", "en"), "Copy selected table as XML")
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
            ):
                key = f"table.error.{code}"
                self.assertNotEqual(t(key, language), key)


if __name__ == "__main__":
    unittest.main()
