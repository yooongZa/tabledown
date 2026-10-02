"""Preserve cell text when generated Markdown is copied back to a table."""

from html import escape
from string import punctuation
import unittest

from bs4 import BeautifulSoup

from tablemark.converter.html_to_md import html_table_to_markdown
from tablemark.converter.md_to_tsv import (
    is_markdown_table,
    markdown_table_to_html,
    markdown_table_to_rows,
)


def _table_with_cells(cells):
    return "| 값 |\n| --- |\n" + "\n".join(f"| {cell} |" for cell in cells)


class MarkdownRoundtripTests(unittest.TestCase):
    def test_automatic_markdown_roundtrip_preserves_literals_in_multiple_columns(self):
        values = [
            r"x\|y", r"\*", r"\\", r"C:\Temp\file", "x|y", "ends|",
            "&copy;", "&lt;", "&amp;lt;", "&#124;", "<br>", "<BR />",
            "첫째\n둘째", "**표시**", "`코드`", "[항목]", "한글 🍎", punctuation,
        ]
        html = "<table><tr><th>값</th><th>확인</th></tr>" + "".join(
            "<tr><td>" + escape(value).replace("\n", "<br>")
            + "</td><td>tail</td></tr>" for value in values
        ) + "</table>"
        markdown = html_table_to_markdown(html)
        expected = [["값", "확인"], *[[value, "tail"] for value in values]]

        self.assertTrue(is_markdown_table(markdown))
        self.assertEqual(markdown_table_to_rows(markdown), expected)
        soup = BeautifulSoup(markdown_table_to_html(markdown), "html.parser")
        for br in soup.find_all("br"):
            br.replace_with("\n")
        self.assertEqual([
            [cell.get_text() for cell in row.find_all(["th", "td"])]
            for row in soup.find_all("tr")
        ], expected)

    def test_automatic_markdown_backslashes_before_pipes_keep_header_and_body_width(self):
        for count in range(6):
            with self.subTest(backslashes=count):
                value = "left" + "\\" * count + "|right"
                html = (
                    "<table><tr><th>" + escape(value) + "</th><th>확인</th></tr>"
                    "<tr><td>" + escape(value) + "</td><td>tail</td></tr></table>"
                )
                markdown = html_table_to_markdown(html)
                self.assertTrue(is_markdown_table(markdown))
                self.assertEqual(markdown_table_to_rows(markdown), [
                    [value, "확인"], [value, "tail"],
                ])

    def test_automatic_markdown_keeps_whitespace_normalization_and_empty_column_trimming(self):
        html = (
            "<table><tr><th> 値 </th><th>確認</th><th></th></tr>"
            "<tr><td> \r\n  A\t  B \r\n\r\n C  \r\n </td>"
            "<td> \t </td><td></td></tr></table>"
        )
        markdown = html_table_to_markdown(html)
        self.assertEqual(markdown, "| 値 | 確認 |\n| --- | --- |\n| A B<br><br>C |   |")
        self.assertEqual(markdown_table_to_rows(markdown), [
            ["値", "確認"], ["A B\n\nC", ""],
        ])

    def test_explicit_markdown_roundtrip_preserves_literal_text_and_line_breaks(self):
        values = [
            "A_B", r"C:\Temp\file", "A&B <태그>", "**표시**", "`코드`",
            "x|y", r"x\|y", "<br>", "첫째\n둘째", "&lt;", "&amp;lt;",
            "&copy;", "1,200원", "12.50%", "한글 🍎", "", punctuation,
        ]
        html = "<table><tr><th>값</th></tr>" + "".join(
            "<tr><td>" + escape(value).replace("\n", "<br>") + "</td></tr>"
            for value in values
        ) + "</table>"
        markdown = html_table_to_markdown(html, preserve_layout=True)

        self.assertTrue(is_markdown_table(markdown))
        self.assertEqual(markdown_table_to_rows(markdown), [["값"], *[[v] for v in values]])
        restored_html = markdown_table_to_html(markdown)
        self.assertTrue(restored_html.startswith('<meta charset="utf-8">'))
        soup = BeautifulSoup(restored_html, "html.parser")
        restored = []
        for cell in soup.find_all("td"):
            for br in cell.find_all("br"):
                br.replace_with("\n")
            restored.append(cell.get_text())
        self.assertEqual(restored, values)

    def test_ascii_punctuation_escapes_decode_once(self):
        markdown = _table_with_cells("\\" + character for character in punctuation)
        self.assertEqual(markdown_table_to_rows(markdown)[1:], [[c] for c in punctuation])

    def test_unknown_backslashes_and_unescaped_markdown_are_preserved(self):
        values = [
            r"C:\Temp\report", r"\n\t\한글", "**강조**", "`A_B`",
            "[문자](주소)", "plain text", "00123", "1.20", "a < b",
        ]
        self.assertEqual(markdown_table_to_rows(_table_with_cells(values))[1:], [[v] for v in values])

    def test_backslash_parity_controls_cell_delimiters(self):
        for count in range(6):
            with self.subTest(backslashes=count):
                escaped = "\\" * count
                suffix = "\\" * (count // 2)
                if count % 2:
                    expected = ["left" + suffix + "|right", "tail"]
                else:
                    expected = ["left" + suffix, "right", "tail"]
                columns = len(expected)
                markdown = (
                    "| " + " | ".join(f"h{i}" for i in range(columns)) + " |\n"
                    "| " + " | ".join("---" for _ in range(columns)) + " |\n"
                    f"| left{escaped}|right | tail |"
                )
                self.assertTrue(is_markdown_table(markdown))
                self.assertEqual(markdown_table_to_rows(markdown)[1], expected)

    def test_header_parity_keeps_strict_column_validation(self):
        for count in range(6):
            with self.subTest(backslashes=count):
                header = "| left" + "\\" * count + "|right | tail |"
                columns = 2 if count % 2 else 3
                correct = "| " + " | ".join("---" for _ in range(columns)) + " |"
                wrong = "| " + " | ".join("---" for _ in range(columns + 1)) + " |"
                self.assertTrue(is_markdown_table(header + "\n" + correct))
                self.assertFalse(is_markdown_table(header + "\n" + wrong))
                self.assertTrue(is_markdown_table(header + "\n" + wrong, strict=False))

    def test_escaped_final_pipe_is_cell_data_without_outer_delimiters(self):
        markdown = "first | second\\|\n--- | ---\nvalue | final\\|"
        self.assertTrue(is_markdown_table(markdown))
        self.assertEqual(markdown_table_to_rows(markdown), [
            ["first", "second|"], ["value", "final|"],
        ])

    def test_entities_and_html_breaks_are_decoded_only_once(self):
        cases = (
            ("&amp;lt;", "&lt;"),
            ("&amp;amp;lt;", "&amp;lt;"),
            ("&lt;br&gt;", "<br>"),
            ("&#60;br&#62;", "<br>"),
            ("&#x1F34E;", "🍎"),
            ("&bogus;", "&bogus;"),
            ("A&amp B", "A&amp B"),
            (r"\&lt;br&gt;", "&lt;br>"),
            (r"\<br>", "<br>"),
            ("first<br>second<BR />third", "first\nsecond\nthird"),
        )
        markdown = _table_with_cells(source for source, _ in cases)
        self.assertEqual(markdown_table_to_rows(markdown)[1:], [[value] for _, value in cases])

    def test_generated_html_distinguishes_literal_break_from_a_real_newline(self):
        markdown = "| 값 |\n| --- |\n| first<br>second |\n| &lt;br&gt; |\n| &amp;lt; |"
        self.assertEqual(markdown_table_to_html(markdown), (
            '<meta charset="utf-8"><table><tr><th>값</th></tr>'
            '<tr><td>first<br>second</td></tr><tr><td>&lt;br&gt;</td></tr>'
            '<tr><td>&amp;lt;</td></tr></table>'
        ))

    def test_plain_tables_keep_source_text_and_separator_shaped_data(self):
        markdown = "| 항목 | 값 |\n| --- | --- |\n| 사과 | 00123 |\n| - | -- |"
        original = markdown
        self.assertEqual(markdown_table_to_rows(markdown), [
            ["항목", "값"], ["사과", "00123"], ["-", "--"],
        ])
        self.assertEqual(markdown, original)
        self.assertEqual(markdown_table_to_html(markdown), (
            '<meta charset="utf-8"><table><tr><th>항목</th><th>값</th></tr>'
            '<tr><td>사과</td><td>00123</td></tr><tr><td>-</td><td>--</td></tr></table>'
        ))


if __name__ == "__main__":
    unittest.main()
