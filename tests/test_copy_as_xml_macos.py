"""macOS manual table-to-XML clipboard action regression tests."""

from __future__ import annotations

import threading
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from tablemark.app import TabledownApp
from tablemark.i18n import t


HTML_TABLE = (
    "<table><tr><th>상품</th><th>값</th></tr>"
    "<tr><td>사과</td><td>2</td></tr></table>"
)
HTML_MULTIPLE_TABLES = HTML_TABLE + HTML_TABLE


class CopyAsXmlActionTests(unittest.TestCase):
    @staticmethod
    def _app():
        return SimpleNamespace(
            fill_blanks=False,
            lang="ko",
            _clipboard_operation_lock=threading.Lock(),
            _clipboard_table_model=TabledownApp._clipboard_table_model,
            _flash_icon_success=Mock(),
        )

    def test_success_replaces_every_clipboard_format_with_xml_text(self):
        app = self._app()
        with (
            patch("tablemark.app.read_clipboard", return_value={"html": HTML_TABLE}),
            patch("tablemark.app.write_text_only_clipboard") as text_only_writer,
            patch("tablemark.app.write_clipboard") as preserving_writer,
            patch("tablemark.app.rumps.alert") as alert,
            patch("tablemark.app.log"),
        ):
            TabledownApp.copy_as_xml(app, None)

        xml = text_only_writer.call_args.args[0]
        self.assertTrue(xml.startswith("<표>"))
        self.assertIn('<열 n="상품">사과</열>', xml)
        text_only_writer.assert_called_once_with(xml, mark_generated=True)
        preserving_writer.assert_not_called()
        alert.assert_not_called()
        app._flash_icon_success.assert_called_once_with()

    def test_no_table_leaves_clipboard_untouched_and_shows_guidance(self):
        app = self._app()
        with (
            patch("tablemark.app.read_clipboard", return_value={"text": "일반 문장"}),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.rumps.alert") as alert,
            patch("tablemark.app.log"),
        ):
            TabledownApp.copy_as_xml(app, None)

        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()
        alert.assert_called_once_with(
            title=t("xml.no_table_title", "ko"),
            message=t("xml.no_table_message", "ko"),
        )

    def test_multiple_tables_leave_clipboard_untouched_and_show_specific_guidance(self):
        app = self._app()
        with (
            patch(
                "tablemark.app.read_clipboard",
                return_value={"html": HTML_MULTIPLE_TABLES},
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.rumps.alert") as alert,
            patch("tablemark.app.log") as logger,
        ):
            TabledownApp.copy_as_xml(app, None)

        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with("copy as xml failed: multiple_tables")
        alert.assert_called_once_with(
            title=t("xml.no_table_title", "ko"),
            message=t("xml.multiple_tables_message", "ko"),
        )

    def test_writer_failure_does_not_report_success(self):
        app = self._app()
        with (
            patch("tablemark.app.read_clipboard", return_value={"html": HTML_TABLE}),
            patch(
                "tablemark.app.write_text_only_clipboard",
                side_effect=OSError("PRIVATE_CLIPBOARD_DETAIL"),
            ),
            patch(
                "tablemark.app.write_clipboard",
                side_effect=OSError("PRIVATE_CLIPBOARD_DETAIL"),
            ),
            patch("tablemark.app.rumps.alert") as alert,
            patch("tablemark.app.log") as logger,
        ):
            TabledownApp.copy_as_xml(app, None)

        app._flash_icon_success.assert_not_called()
        logger.assert_called_once_with("copy as xml failed: OSError")
        alert.assert_called_once_with(
            title=t("xml.no_table_title", "ko"),
            message=t("xml.no_table_message", "ko"),
        )


if __name__ == "__main__":
    unittest.main()
