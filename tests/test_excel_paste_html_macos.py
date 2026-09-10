"""Excel paste HTML compatibility without accessing the system clipboard."""

from __future__ import annotations

import re
import sys
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET

if sys.platform == "darwin":
    from AppKit import (
        NSAttributedString,
        NSDocumentTypeDocumentAttribute,
        NSHTMLTextDocumentType,
        NSPasteboard,
    )
    from bs4 import BeautifulSoup

    from tablemark.app import MAX_MARKDOWN_EXPORT_BYTES, _EXPORT_MARKDOWN, _EXPORT_TABLE
    from tablemark.clipboard import (
        GENERATED_MARKER_TYPES, HTML_TYPES, STRING_TYPES, write_table_clipboard,
    )
    from tablemark.converter.html_to_md import html_table_to_markdown
    from tablemark.converter.md_to_tsv import markdown_table_to_rows
    from tablemark.converter.table_xml import TableXmlTooLargeError, model_to_xml
    from tablemark.excel_table import (
        excel_table_selection_to_html,
        excel_table_selection_to_model,
    )
    from tests.test_copy_as_markdown_macos import _selection
    from tests import test_copy_as_xml_macos as xml_tests


@unittest.skipUnless(sys.platform == "darwin", "requires macOS HTML importer")
class ExcelPasteHtmlTests(unittest.TestCase):
    def setUp(self):
        # Any unexpected production clipboard route fails before native access.
        patcher = patch("tablemark.clipboard.NSPasteboard")
        self.boards = patcher.start()
        self.addCleanup(patcher.stop)
        self.boards.generalPasteboard.side_effect = AssertionError(
            "system clipboard access is forbidden in these tests"
        )

    def _capture(self, selection, *, xml=False):
        app = xml_tests.CopyAsXmlActionTests._app()
        with (
            patch("tablemark.app.read_stable_selected_excel_table",
                  return_value=selection) as reader,
            patch("tablemark.app.write_table_clipboard") as table_writer,
            patch("tablemark.app.write_text_only_clipboard") as xml_writer,
            patch("tablemark.app.write_clipboard") as preserving_writer,
            patch("tablemark.app.read_clipboard") as clipboard_reader,
            patch("tablemark.app.log"),
        ):
            outcome = app._perform_explicit_export(
                _EXPORT_TABLE if xml else _EXPORT_MARKDOWN, 17, False,
            )
        self.assertEqual(outcome, ("success", None))
        reader.assert_called_once_with()
        preserving_writer.assert_not_called()
        clipboard_reader.assert_not_called()
        if xml:
            table_writer.assert_not_called()
            xml_writer.assert_called_once()
            return xml_writer.call_args.args[0]
        xml_writer.assert_not_called()
        table_writer.assert_called_once()
        return table_writer.call_args.args

    @staticmethod
    def _soup(html):
        # BeautifulSoup otherwise normalizes whitespace-only text nodes. This
        # option reads original HTML characters; AppKit is tested independently.
        return BeautifulSoup(
            html, "html.parser", preserve_whitespace_tags={"td", "span"},
        )

    def _cell_values(self, html):
        soup = self._soup(html)
        values = []
        for cell in soup.find("table").find_all("td"):
            for br in cell.find_all("br"):
                br.replace_with("\n")
            values.append(cell.get_text())
        return values

    def test_default_structural_html_and_markdown_remain_identical(self):
        selected = _selection([
            ["공백", "기호", None],
            ["  사과   🍎  ", "A&B <br>\n끝", None],
        ])
        legacy_html = (
            "<table><tr><td>공백</td><td>기호</td><td></td></tr>"
            "<tr><td>  사과   🍎  </td>"
            "<td>A&amp;B &lt;br&gt;<br>끝</td><td></td></tr></table>"
        )
        self.assertEqual(excel_table_selection_to_html(selected), legacy_html)
        self.assertEqual(
            excel_table_selection_to_html(selected, for_excel_paste=False),
            legacy_html,
        )
        default_md = html_table_to_markdown(legacy_html)
        layout_md = html_table_to_markdown(legacy_html, preserve_layout=True)
        excel_table_selection_to_html(selected, for_excel_paste=True)
        after = excel_table_selection_to_html(selected)
        self.assertEqual(after, legacy_html)
        self.assertEqual(html_table_to_markdown(after), default_md)
        self.assertEqual(
            html_table_to_markdown(after, preserve_layout=True), layout_md,
        )
        self.assertEqual(markdown_table_to_rows(layout_md), [
            ["공백", "기호", ""],
            ["사과   🍎", "A&B <br>\n끝", ""],
        ])
        # Markdown text keeps source spaces; its existing TSV parser
        # trims cell-edge padding and is outside this HTML change.
        self.assertIn("|   사과   🍎   |", layout_md)
        self.assertNotIn("mso", layout_md)

    def test_general_xml_values_and_source_metadata_do_not_change(self):
        selected = _selection([
            ["공백", "기호", "빈열"],
            ["  ACME   Inc.  ", "A\u00a0B & <br>\n끝", None],
        ])
        before_model = excel_table_selection_to_model(selected)
        before_xml = self._capture(selected, xml=True)
        self._capture(selected)
        self.assertEqual(excel_table_selection_to_model(selected), before_model)
        self.assertEqual(self._capture(selected, xml=True), before_xml)
        root = ET.fromstring(before_xml)
        self.assertEqual(root.attrib["주소"], "$A$1:$C$2")
        self.assertEqual(root.attrib["열수"], "3")
        self.assertEqual(root.attrib["병합범위"], "")
        self.assertEqual(
            [cell.text or "" for cell in root.findall(".//열")],
            ["  ACME   Inc.  ", "A\u00a0B & <br>\n끝", ""],
        )
        self.assertNotIn("mso", model_to_xml(*before_model))

    def test_explicit_markdown_uses_structural_and_paste_html_separately(self):
        selected = _selection([["항목", "값"], ["  A   B  ", "첫째\n둘째"]])
        with patch(
            "tablemark.app.excel_table_selection_to_html",
            wraps=excel_table_selection_to_html,
        ) as bridge:
            markdown, html = self._capture(selected)
        self.assertEqual(len(bridge.call_args_list), 2)
        self.assertEqual(bridge.call_args_list[0].args, (selected,))
        self.assertEqual(bridge.call_args_list[0].kwargs, {})
        self.assertEqual(bridge.call_args_list[1].args, (selected,))
        paste_options = bridge.call_args_list[1].kwargs
        self.assertEqual(set(paste_options), {"for_excel_paste", "max_output_bytes"})
        self.assertIs(paste_options["for_excel_paste"], True)
        table_start = html.index("<table>")
        table_end = html.index("</table>") + len("</table>")
        wrapper_bytes = len((html[:table_start] + html[table_end:]).encode("utf-8"))
        self.assertEqual(
            paste_options["max_output_bytes"],
            MAX_MARKDOWN_EXPORT_BYTES - len(markdown.encode("utf-8")) - wrapper_bytes,
        )
        self.assertTrue(markdown.startswith("\n|"))
        self.assertTrue(markdown.endswith("\n"))
        self.assertNotIn("mso", markdown)
        self.assertEqual(markdown_table_to_rows(markdown),
                         [["항목", "값"], ["A   B", "첫째\n둘째"]])
        self.assertIn("|   A   B   |", markdown)
        soup = self._soup(html)
        self.assertEqual(soup.find("meta")["charset"], "utf-8")
        self.assertIn("white-space:pre-wrap", soup.find("style").get_text())
        self.assertIn("mso-data-placement:same-cell", html)

    def test_explicit_combined_budget_includes_markdown_and_html_wrapper(self):
        selected = _selection([["항목", "값"], ["  A B  ", "한글🙂\n끝"]])
        markdown, html = self._capture(selected)
        exact_size = len(markdown.encode("utf-8")) + len(html.encode("utf-8"))
        with patch("tablemark.app.MAX_MARKDOWN_EXPORT_BYTES", exact_size):
            self.assertEqual(self._capture(selected), (markdown, html))
        app = xml_tests.CopyAsXmlActionTests._app()
        with (
            patch("tablemark.app.MAX_MARKDOWN_EXPORT_BYTES", exact_size - 1),
            patch("tablemark.app.read_stable_selected_excel_table", return_value=selected),
            patch("tablemark.app.write_table_clipboard") as writer,
            patch("tablemark.app.log"),
        ):
            outcome = app._perform_explicit_export(_EXPORT_MARKDOWN, 17, False)
        self.assertEqual(outcome, ("error", "output_too_large"))
        writer.assert_not_called()
        app._flash_icon_success.assert_not_called()

    def test_non_office_text_keeps_ascii_spaces_original_nbsp_and_escapes(self):
        values = [
            "  A   B  ", "A\u00a0B", "\u00a0", " ",
            'A&B <br> <br/> "quoted" \\ | &#160; &nbsp;',
            "<!--[if mso]><b>literal</b><![endif]-->",
        ]
        selected = _selection([values])
        html = excel_table_selection_to_html(selected, for_excel_paste=True)
        self.assertEqual(self._cell_values(html), values)
        self.assertEqual(sum(value.count("\u00a0") for value in self._cell_values(html)), 2)
        self.assertNotIn("<b>literal</b>", html)
        self.assertIn("&amp;#160;", html)
        self.assertIn("&amp;nbsp;", html)
        self.assertEqual(len(self._soup(html).find_all("br")), 0)
        office_space_runs = re.findall(
            r'<!--\[if mso\]><span style="mso-spacerun:yes">(.*?)</span><!\[endif\]-->',
            html,
        )
        expected_lengths = [len(run) for value in values for run in re.findall(r" +", value)]
        self.assertEqual(office_space_runs, ["&#160;" * n for n in expected_lengths])

    def test_newlines_stay_in_cells_and_literal_break_tags_remain_text(self):
        source = "앞\r\n중간\r뒤\n\n끝 <br> <br/>"
        selected = _selection([["줄바꿈", None], [source, None]])
        html = excel_table_selection_to_html(selected, for_excel_paste=True)
        soup = self._soup(html)
        rows = soup.find("table").find_all("tr")
        self.assertEqual([len(row.find_all("td")) for row in rows], [2, 2])
        self.assertEqual(len(soup.find_all("br")), 4)
        self.assertTrue(all(
            br.get("style") == "mso-data-placement:same-cell"
            for br in soup.find_all("br")
        ))
        self.assertEqual(self._cell_values(html),
                         ["줄바꿈", "", "앞\n중간\n뒤\n\n끝 <br> <br/>", ""])

    def test_merges_empty_rows_and_selected_trailing_columns_are_preserved(self):
        selected = _selection([
            ["  제목  ", None, None, None],
            [None, None, None, None],
            ["  항목  ", "A\u00a0B", "첫째\n둘째", None],
        ], merges=("$A$1:$C$2",))
        baseline = self._soup(excel_table_selection_to_html(selected))
        html = excel_table_selection_to_html(selected, for_excel_paste=True)
        soup = self._soup(html)
        rows = soup.find("table").find_all("tr")
        self.assertEqual([len(row.find_all("td")) for row in rows], [2, 1, 4])
        self.assertEqual(
            [cell.attrs for cell in soup.find_all("td")],
            [cell.attrs for cell in baseline.find_all("td")],
        )
        title = soup.find("td", rowspan="2", colspan="3")
        self.assertIsNotNone(title)
        self.assertEqual(title.get_text(), "  제목  ")
        self.assertEqual(self._cell_values(html),
                         ["  제목  ", "", "", "  항목  ", "A\u00a0B", "첫째\n둘째", ""])
        self.assertEqual(len(rows[1].find_all("td")), 1)
        self.assertEqual(rows[-1].find_all("td")[-1].get_text(), "")

    def test_native_non_office_import_preserves_original_characters(self):
        rows = [
            ["앞", "뒤", "연속", "NBSP", "기호"],
            ["  한글 🍎", "한글  ", "A   B", "A\u00a0B", 'A&B <br> \\ | "인용"'],
            ["한 칸", "여러   칸", "\u00a0", "&#160; &nbsp;", "첫째\n둘째"],
        ]
        markdown, html = self._capture(_selection(rows))
        expected = [line for row in rows for value in row for line in value.splitlines()]
        pasteboard = NSPasteboard.pasteboardWithUniqueName()
        try:
            with patch("tablemark.clipboard.NSPasteboard") as boards:
                boards.generalPasteboard.return_value = pasteboard
                write_table_clipboard(
                    markdown, html,
                    expected_change_count=int(pasteboard.changeCount()),
                )
            self.assertEqual(
                {str(kind) for kind in pasteboard.types()},
                STRING_TYPES | HTML_TYPES | set(GENERATED_MARKER_TYPES),
            )
            for kind in STRING_TYPES:
                self.assertEqual(str(pasteboard.stringForType_(kind)), markdown)
            for kind in sorted(HTML_TYPES):
                with self.subTest(html_type=kind):
                    self.assertEqual(str(pasteboard.stringForType_(kind)), html)
                    value, _, error = (
                        NSAttributedString.alloc()
                        .initWithData_options_documentAttributes_error_(
                            pasteboard.dataForType_(kind),
                            {NSDocumentTypeDocumentAttribute: NSHTMLTextDocumentType},
                            None, None,
                        )
                    )
                    self.assertIsNone(error)
                    self.assertIsNotNone(value)
                    actual = str(value.string()).splitlines()
                    self.assertEqual(actual, expected)
                    self.assertEqual(sum(v.count("\u00a0") for v in actual), 2)
        finally:
            pasteboard.releaseGlobally()
        self.boards.generalPasteboard.assert_not_called()


@unittest.skipUnless(sys.platform == "darwin", "requires macOS reader types")
class ExcelPasteHtmlBudgetTests(unittest.TestCase):
    def test_exact_utf8_budget_preserves_small_and_unicode_output(self):
        selected = _selection([
            ["한글🙂", "  &<> \" 인용  ", None],
            ["앞\r\n뒤\r끝\n", "A\u00a0B &#160;", " "],
        ])
        for for_excel_paste in (False, True):
            with self.subTest(for_excel_paste=for_excel_paste):
                expected = excel_table_selection_to_html(
                    selected, for_excel_paste=for_excel_paste,
                )
                exact_bytes = len(expected.encode("utf-8"))
                self.assertEqual(excel_table_selection_to_html(
                    selected, for_excel_paste=for_excel_paste,
                    max_output_bytes=exact_bytes,
                ), expected)
                self.assertEqual(excel_table_selection_to_html(
                    selected, for_excel_paste=for_excel_paste,
                    max_output_bytes=exact_bytes + 1,
                ), expected)
                with self.assertRaises(TableXmlTooLargeError):
                    excel_table_selection_to_html(
                        selected, for_excel_paste=for_excel_paste,
                        max_output_bytes=exact_bytes - 1,
                    )

    def test_negative_budget_fails_before_selection_validation_or_expansion(self):
        selected = _selection([["value"]])
        with (
            patch("tablemark.excel_table._rectangle_bounds") as bounds,
            patch("tablemark.excel_table.escape") as escape_cell,
            self.assertRaises(TableXmlTooLargeError),
        ):
            excel_table_selection_to_html(
                selected, for_excel_paste=True, max_output_bytes=-1,
            )
        bounds.assert_not_called()
        escape_cell.assert_not_called()

    def test_empty_cell_markup_is_part_of_budget(self):
        selected = _selection([[None]])
        for for_excel_paste in (False, True):
            expected = (
                "<table><!--StartFragment--><tr><td></td></tr><!--EndFragment--></table>"
                if for_excel_paste else "<table><tr><td></td></tr></table>"
            )
            with self.subTest(for_excel_paste=for_excel_paste):
                self.assertEqual(excel_table_selection_to_html(
                    selected, for_excel_paste=for_excel_paste,
                    max_output_bytes=len(expected),
                ), expected)
                for budget in (0, len(expected) - 1):
                    with self.assertRaises(TableXmlTooLargeError):
                        excel_table_selection_to_html(
                            selected, for_excel_paste=for_excel_paste,
                            max_output_bytes=budget,
                        )

    def test_single_large_cell_fails_before_escape_or_office_expansion(self):
        # This is valid Excel cell text. Its 20,000 input characters used to
        # allocate 1,190,033 bytes of HTML before the caller could reject it.
        selected = _selection([["x " * 10_000]])
        with (
            patch("tablemark.excel_table.escape") as escape_cell,
            patch("tablemark.excel_table.re.sub") as expand_spaces,
            self.assertRaises(TableXmlTooLargeError),
        ):
            excel_table_selection_to_html(
                selected, for_excel_paste=True, max_output_bytes=1_000_000,
            )
        escape_cell.assert_not_called()
        expand_spaces.assert_not_called()

    def test_large_entity_expansion_fails_before_escaping(self):
        selected = _selection([["&" * 20_000]])
        with (
            patch("tablemark.excel_table.escape") as escape_cell,
            self.assertRaises(TableXmlTooLargeError),
        ):
            excel_table_selection_to_html(
                selected, for_excel_paste=True, max_output_bytes=30_000,
            )
        escape_cell.assert_not_called()

    def test_multiple_cells_share_budget_and_reject_before_next_expansion(self):
        selected = _selection([["A B", "C D"]])
        expected = excel_table_selection_to_html(selected, for_excel_paste=True)
        with (
            patch("tablemark.excel_table.re.sub", wraps=re.sub) as expand_spaces,
            self.assertRaises(TableXmlTooLargeError),
        ):
            excel_table_selection_to_html(
                selected, for_excel_paste=True,
                max_output_bytes=len(expected.encode("utf-8")) - 1,
            )
        self.assertEqual(expand_spaces.call_count, 1)

    def test_excel_fragment_markers_are_unique_inside_table_and_budgeted(self):
        selected = _selection([["literal <!--StartFragment-->", "값"]])
        default_html = excel_table_selection_to_html(selected)
        self.assertNotIn("<!--StartFragment-->", default_html)
        self.assertNotIn("<!--EndFragment-->", default_html)
        excel_html = excel_table_selection_to_html(selected, for_excel_paste=True)
        self.assertTrue(excel_html.startswith("<table><!--StartFragment-->"))
        self.assertTrue(excel_html.endswith("<!--EndFragment--></table>"))
        self.assertEqual(excel_html.count("<!--StartFragment-->"), 1)
        self.assertEqual(excel_html.count("<!--EndFragment-->"), 1)
        exact_bytes = len(excel_html.encode("utf-8"))
        self.assertEqual(excel_table_selection_to_html(
            selected, for_excel_paste=True, max_output_bytes=exact_bytes,
        ), excel_html)
        with self.assertRaises(TableXmlTooLargeError):
            excel_table_selection_to_html(
                selected, for_excel_paste=True, max_output_bytes=exact_bytes - 1,
            )

    def test_merge_attributes_covered_rows_and_empty_columns_count_exactly(self):
        selected = _selection([
            ["  제목  ", None, None],
            [None, None, None],
            ["끝", "줄\n바꿈", None],
        ], merges=("$A$1:$B$2",))
        expected = excel_table_selection_to_html(selected, for_excel_paste=True)
        size = len(expected.encode("utf-8"))
        self.assertEqual(excel_table_selection_to_html(
            selected, for_excel_paste=True, max_output_bytes=size,
        ), expected)
        with self.assertRaises(TableXmlTooLargeError):
            excel_table_selection_to_html(
                selected, for_excel_paste=True, max_output_bytes=size - 1,
            )

    def test_failed_budget_does_not_change_later_calls_or_default_xml_model(self):
        selected = _selection([["항목", "값"], ["  A  ", "B\nC"]])
        default_html = excel_table_selection_to_html(selected)
        default_model = excel_table_selection_to_model(selected)
        expected = excel_table_selection_to_html(selected, for_excel_paste=True)
        with self.assertRaises(TableXmlTooLargeError):
            excel_table_selection_to_html(
                selected, for_excel_paste=True, max_output_bytes=0,
            )
        self.assertEqual(excel_table_selection_to_html(
            selected, for_excel_paste=True,
            max_output_bytes=len(expected.encode("utf-8")),
        ), expected)
        self.assertEqual(excel_table_selection_to_html(selected), default_html)
        self.assertEqual(excel_table_selection_to_model(selected), default_model)


if __name__ == "__main__":
    unittest.main()
