"""Direct Excel selection Markdown copying without the system clipboard."""

from __future__ import annotations

import sys
import unittest
from unittest.mock import patch

if sys.platform == "darwin":
    from bs4 import BeautifulSoup

    from tablemark.app import TabledownApp
    from tablemark.clipboard import ClipboardChangedError, ClipboardWriteError
    from tablemark.converter.html_to_md import html_table_to_markdown
    from tablemark.converter.md_to_tsv import markdown_table_to_rows
    from tablemark.excel_formula import (
        DISPLAY_OVERFLOW,
        EXCEL_NOT_RUNNING,
        MULTIPLE_AREAS,
        PARTIAL_MERGE,
        SELECTION_CHANGED,
        ExcelFormulaError,
    )
    from tablemark.excel_table import (
        TABLE_RESULT_STATUS,
        excel_table_selection_to_html,
        parse_excel_table_result,
    )
    from tablemark.i18n import t
    from tests import test_copy_as_xml_macos as xml_tests


def _selection(rows, merges=()):
    """Use the production Excel payload validator for synthetic displayed values."""
    row_count = len(rows)
    column_count = len(rows[0])
    values = [value for row in rows for value in row]
    address = (
        "$A$1" if row_count == column_count == 1
        else f"$A$1:${chr(64 + column_count)}${row_count}"
    )
    encoded = [
        ("" if value is None else value)
        .replace("\ue001", "\ue001e").replace("\ue000", "\ue001d")
        for value in values
    ]
    return parse_excel_table_result([
        TABLE_RESULT_STATUS, "Synthetic.xlsx", "매출", address,
        str(row_count), str(column_count),
        "".join("b" if value is None else "v" for value in values),
        "\ue000".join(encoded), list(merges),
    ])


def _merged_selection():
    return _selection([
        ["직급", "1분기", None, "메모"],
        [None, "1월", "2월", "상태"],
        ["A팀", "10", None, "자료"],
        [None, "20", "30", None],
    ], ("$A$1:$A$2", "$B$1:$C$1", "$A$3:$A$4"))


@unittest.skipUnless(sys.platform == "darwin", "requires macOS app dependencies")
class CopyAsMarkdownActionTests(unittest.TestCase):
    def setUp(self):
        # Even an accidentally unpatched route must not access the real clipboard.
        pasteboards = self._patch("tablemark.clipboard.NSPasteboard")
        pasteboards.generalPasteboard.side_effect = AssertionError(
            "system clipboard access is forbidden in these tests"
        )
        self.writer = self._patch("tablemark.app.write_table_clipboard")
        self.old_writer = self._patch("tablemark.app.write_clipboard")
        self.xml_writer = self._patch("tablemark.app.write_text_only_clipboard")
        self.clipboard_reader = self._patch("tablemark.app.read_clipboard")
        self.formula_reader = self._patch("tablemark.app.read_stable_selected_excel_formulas")
        self.reader = self._patch(
            "tablemark.app.read_stable_selected_excel_table",
            return_value=_selection([["항목", "값"], ["사과", "2"]]),
        )
        self.logger = self._patch("tablemark.app.log")

    def tearDown(self):
        self.old_writer.assert_not_called()
        self.xml_writer.assert_not_called()
        self.clipboard_reader.assert_not_called()
        self.formula_reader.assert_not_called()

    def _patch(self, target, **kwargs):
        patcher = patch(target, **kwargs)
        result = patcher.start()
        self.addCleanup(patcher.stop)
        return result

    @staticmethod
    def _app():
        return xml_tests.CopyAsXmlActionTests._app()

    @staticmethod
    def _run(app):
        with (
            patch("tablemark.app.clipboard_change_count", return_value=17),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=lambda callback, *args: callback(*args),
            ),
        ):
            app.copy_as_markdown(None)

    def _copied(self):
        self.writer.assert_called_once()
        self.assertEqual(self.writer.call_args.kwargs, {
            "mark_generated": True, "expected_change_count": 17,
        })
        return self.writer.call_args.args

    def test_success_uses_one_snapshot_and_preserves_selected_values_and_columns(self):
        selected = _selection([
            ["항목", "금액", "설명", "날짜", "비율", None],
            ["  사과 🍎 | A&B  ", "₩1,200", "첫째\r\n둘째\n<br>\n\n끝",
             "2026-09-10", "12.50%", None],
        ])
        self.reader.return_value = selected
        app = self._app()
        self._run(app)
        markdown, html = self._copied()
        self.reader.assert_called_once_with()
        self.assertEqual(markdown.strip("\n"), (
            "| 항목 | 금액 | 설명 | 날짜 | 비율 |   |\n"
            "| --- | --- | --- | --- | --- | --- |\n"
            "|   사과 🍎 \\| A&amp;B   | ₩1,200 | "
            "첫째<br>둘째<br>&lt;br&gt;<br><br>끝 | 2026-09-10 | 12.50% |   |"
        ))
        self.assertTrue(markdown.startswith("\n|"))
        # Whitespace-only spans are intentional source data. The default
        # BeautifulSoup parser normalizes those nodes before get_text().
        soup = BeautifulSoup(
            html, "html.parser", preserve_whitespace_tags={"td", "span"},
        )
        self.assertEqual(soup.find("meta")["charset"], "utf-8")
        self.assertIn("white-space:pre-wrap", soup.find("style").get_text())
        html_rows = soup.find("table").find_all("tr")
        self.assertEqual([len(row.find_all("td")) for row in html_rows], [6, 6])
        actual_values = []
        for cell in soup.find("table").find_all("td"):
            for br in cell.find_all("br"):
                br.replace_with("\n")
            actual_values.append(cell.get_text())
        self.assertEqual(actual_values, [
            "" if value is None else value.replace("\r\n", "\n")
            for value in selected.values
        ])
        app._flash_icon_success.assert_called_once_with()
        app._safe_alert.assert_not_called()
        self.assertFalse(app._explicit_export_lock.locked())
        self.assertIsNone(app._explicit_export_active)

    def test_single_cell_and_single_row_can_be_copied(self):
        for rows in ([["한글 🍎"]], [["항목", "금액", None]]):
            with self.subTest(columns=len(rows[0])):
                self.writer.reset_mock()
                self.reader.return_value = _selection(rows)
                app = self._app()
                self._run(app)
                markdown, html = self._copied()
                self.assertEqual(markdown_table_to_rows(markdown), [
                    ["" if value is None else value for value in row] for row in rows
                ])
                self.assertEqual(len(BeautifulSoup(html, "html.parser").find_all("td")), len(rows[0]))
                app._safe_alert.assert_not_called()
                app._flash_icon_success.assert_called_once_with()

    def test_merged_headers_follow_fill_option_and_leave_data_blanks(self):
        selected = _merged_selection()
        self.reader.return_value = selected
        html_outputs = []
        for fill_blanks in (False, True):
            with self.subTest(fill_blanks=fill_blanks):
                self.writer.reset_mock()
                app = self._app()
                app.fill_blanks = fill_blanks
                self._run(app)
                markdown, html = self._copied()
                html_outputs.append(html)
                rows = markdown_table_to_rows(markdown)
                self.assertEqual(rows[0], ["직급", "1분기", "1분기" if fill_blanks else "", "메모"])
                self.assertEqual(rows[1], ["직급" if fill_blanks else "", "1월", "2월", "상태"])
                self.assertEqual(rows[2], ["A팀", "10", "", "자료"])
                self.assertEqual(rows[3], ["A팀" if fill_blanks else "", "20", "30", ""])
                soup = BeautifulSoup(html, "html.parser")
                self.assertEqual(len(soup.find_all("td", rowspan="2")), 2)
                self.assertEqual(len(soup.find_all("td", colspan="2")), 1)
        self.assertEqual(html_outputs[0], html_outputs[1])

    def test_full_width_multirow_title_keeps_four_selected_rows_without_changing_automatic_output(self):
        selected = _selection([
            ["2026 매출", None], [None, None],
            ["품목", "금액"], ["사과", "1,200"],
        ], ("$A$1:$B$2",))
        self.reader.return_value = selected
        source_html = excel_table_selection_to_html(selected)
        automatic_before = (
            "| 2026 매출 |   |\n| --- | --- |\n"
            "| 품목 | 금액 |\n| 사과 | 1,200 |"
        )
        for fill_blanks in (False, True):
            with self.subTest(fill_blanks=fill_blanks):
                self.writer.reset_mock()
                app = self._app()
                app.fill_blanks = fill_blanks
                self._run(app)
                markdown, html = self._copied()
                rows = markdown_table_to_rows(markdown)
                self.assertEqual(len(rows), 4)
                self.assertEqual([len(row) for row in rows], [2, 2, 2, 2])
                self.assertEqual(rows[0], ["2026 매출", ""])
                self.assertEqual(rows[2], ["품목", "금액"])
                self.assertEqual(rows[3], ["사과", "1,200"])
                if not fill_blanks:
                    self.assertEqual(rows[1], ["", ""])
                soup = BeautifulSoup(html, "html.parser")
                self.assertEqual(len(soup.find("table").find_all("tr")), 4)
                title = soup.find("td", rowspan="2", colspan="2")
                self.assertIsNotNone(title)
                self.assertEqual(title.get_text(), "2026 매출")
                self.assertEqual(
                    html_table_to_markdown(source_html, fill_merged_headers=fill_blanks),
                    automatic_before,
                )

    def test_unmerged_group_blanks_fill_only_before_populated_data_column(self):
        self.reader.return_value = _selection([
            ["그룹", "항목", "값"], ["부장", "사과", None], [None, "배", "2"],
        ])
        app = self._app()
        app.fill_blanks = True
        self._run(app)
        markdown, _ = self._copied()
        self.assertEqual(markdown_table_to_rows(markdown), [
            ["그룹", "항목", "값"], ["부장", "사과", ""], ["부장", "배", "2"],
        ])

    def test_automatic_converter_defaults_remain_unchanged_after_layout_export(self):
        html = excel_table_selection_to_html(_selection([
            ["항목", "설명", None], ["  사과   🍎  ", "첫째\n<br>", None],
        ]))
        expected = (
            "| 항목 | 설명 |\n| --- | --- |\n"
            "| 사과 🍎 | 첫째<br><br> |"
        )
        self.assertEqual(html_table_to_markdown(html), expected)
        selected = html_table_to_markdown(html, preserve_layout=True)
        self.assertIn("&lt;br&gt;", selected)
        self.assertIn("  사과   🍎  ", selected)
        self.assertEqual(len(markdown_table_to_rows(selected)[0]), 3)
        self.assertEqual(html_table_to_markdown(html), expected)

    def test_output_limit_counts_both_formats_in_utf8_and_accepts_exact_boundary(self):
        self.reader.return_value = _selection([["항목", "값"], ["한글 🍎", "₩1,200"]])
        self._run(self._app())
        markdown, html = self._copied()
        combined_bytes = len(markdown.encode("utf-8")) + len(html.encode("utf-8"))
        self.assertGreater(combined_bytes, len(markdown) + len(html))
        self.assertGreater(combined_bytes - 1, max(len(markdown.encode("utf-8")), len(html.encode("utf-8"))))
        for limit, succeeds in ((combined_bytes, True), (combined_bytes - 1, False)):
            with self.subTest(limit=limit):
                self.writer.reset_mock()
                app = self._app()
                with patch("tablemark.app.MAX_MARKDOWN_EXPORT_BYTES", limit):
                    self._run(app)
                if succeeds:
                    self.writer.assert_called_once()
                    app._flash_icon_success.assert_called_once_with()
                else:
                    self.writer.assert_not_called()
                    app._flash_icon_success.assert_not_called()
                    app._safe_alert.assert_called_once_with("Tabledown", t("markdown.error.output_too_large", "ko"))
                self.assertFalse(app._explicit_export_lock.locked())

    def test_selection_reader_failures_preserve_clipboard_and_release_gate(self):
        for code in (EXCEL_NOT_RUNNING, MULTIPLE_AREAS, PARTIAL_MERGE, DISPLAY_OVERFLOW, SELECTION_CHANGED):
            with self.subTest(code=code):
                self.writer.reset_mock()
                app = self._app()
                self.reader.side_effect = ExcelFormulaError(code)
                self._run(app)
                self.writer.assert_not_called()
                app._flash_icon_success.assert_not_called()
                app._safe_alert.assert_called_once_with("Tabledown", t(f"table.error.{code}", "ko"))
                self.assertFalse(app._explicit_export_lock.locked())
                self.assertIsNone(app._explicit_export_active)

    def test_html_or_markdown_conversion_failure_does_not_write_or_expose_values(self):
        for target in ("excel_table_selection_to_html", "html_table_to_markdown"):
            with self.subTest(target=target):
                app = self._app()
                with patch(f"tablemark.app.{target}", side_effect=ValueError("PRIVATE_VALUE")):
                    self._run(app)
                self.writer.assert_not_called()
                app._flash_icon_success.assert_not_called()
                app._safe_alert.assert_called_once_with("Tabledown", t("table.error.execution_failed", "ko"))
                self.assertNotIn("PRIVATE_VALUE", repr(self.logger.call_args_list))
                self.assertFalse(app._explicit_export_lock.locked())
                self.assertIsNone(app._explicit_export_active)

    def test_new_external_copy_or_failed_writer_does_not_report_success(self):
        for error, code in (
            (ClipboardChangedError("PRIVATE_VALUE"), "clipboard_changed"),
            (ClipboardWriteError("PRIVATE_VALUE"), "clipboard_write_failed"),
        ):
            with self.subTest(code=code):
                self.writer.reset_mock()
                self.writer.side_effect = error
                app = self._app()
                self._run(app)
                self.writer.assert_called_once()
                app._flash_icon_success.assert_not_called()
                app._safe_alert.assert_called_once_with("Tabledown", t(f"markdown.error.{code}", "ko"))
                self.assertNotIn("PRIVATE_VALUE", repr(self.logger.call_args_list))
                self.assertFalse(app._explicit_export_lock.locked())
                self.assertIsNone(app._explicit_export_active)

    def test_stop_before_or_during_read_cancels_without_writing(self):
        selected = self.reader.return_value
        for during_read in (False, True):
            with self.subTest(during_read=during_read):
                self.reader.reset_mock()
                app = self._app()
                if during_read:
                    def stop_after_read():
                        app._stop_watcher.set()
                        return selected
                    self.reader.side_effect = stop_after_read
                else:
                    app._stop_watcher.set()
                self._run(app)
                self.assertEqual(self.reader.call_count, int(during_read))
                self.writer.assert_not_called()
                app._flash_icon_success.assert_not_called()
                app._safe_alert.assert_not_called()
                self.assertFalse(app._explicit_export_lock.locked())

    def test_manual_retry_after_writer_failure_rereads_and_succeeds(self):
        app = self._app()
        self.writer.side_effect = [ClipboardWriteError("clipboard_write_failed"), None]
        self._run(app)
        self.writer.assert_called_once()
        self.assertFalse(app._explicit_export_lock.locked())
        app._flash_icon_success.assert_not_called()
        self._run(app)
        self.assertEqual(self.reader.call_count, 2)
        self.assertEqual(self.writer.call_count, 2)
        app._flash_icon_success.assert_called_once_with()
        app._safe_alert.assert_called_once()
        self.assertFalse(app._explicit_export_lock.locked())
        self.assertIsNone(app._explicit_export_active)

    def test_fill_option_is_captured_before_the_queued_read(self):
        self.reader.return_value = _merged_selection()
        for initial_fill in (False, True):
            with self.subTest(initial_fill=initial_fill):
                self.writer.reset_mock()
                queued = []
                app = self._app()
                app.fill_blanks = initial_fill
                with (
                    patch("tablemark.app.clipboard_change_count", return_value=17),
                    patch("tablemark.app.AppHelper.callAfter", side_effect=lambda cb, *args: queued.append((cb, args))),
                ):
                    app.copy_as_markdown(None)
                app.fill_blanks = not initial_fill
                self.assertEqual(len(queued), 1)
                callback, args = queued[0]
                callback(*args)
                markdown, _ = self._copied()
                self.assertEqual(markdown_table_to_rows(markdown)[3][0], "A팀" if initial_fill else "")

    def test_markdown_action_blocks_other_exports_until_completion(self):
        app = self._app()
        queued = []
        with (
            patch("tablemark.app.clipboard_change_count", return_value=17),
            patch("tablemark.app.AppHelper.callAfter", side_effect=lambda cb, *args: queued.append((cb, args))),
        ):
            app.copy_as_markdown(None)
            for item in (app.copy_markdown_item, app.copy_xml_item, app.copy_excel_formulas_item):
                item._menuitem.setEnabled_.assert_called_with(False)
            app.copy_as_markdown(None)
            app.copy_as_xml(None)
            TabledownApp.copy_selected_excel_formulas(app, None)
        self.assertEqual(len(queued), 1)
        self.assertEqual(app._safe_alert.call_count, 3)
        self.reader.assert_not_called()
        callback, args = queued[0]
        callback(*args)
        self.writer.assert_called_once()
        for item in (app.copy_markdown_item, app.copy_xml_item, app.copy_excel_formulas_item):
            item._menuitem.setEnabled_.assert_called_with(True)
        self.assertFalse(app._explicit_export_lock.locked())


if __name__ == "__main__":
    unittest.main()
