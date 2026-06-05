from __future__ import annotations

import sys
import unittest
from pathlib import Path


WINDOWS_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = WINDOWS_ROOT.parent
for path in (PROJECT_ROOT, WINDOWS_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from tabledown_windows.conversion import WINDOWS_DROP_FORMATS, converted_clipboard
from tabledown_windows.html_clipboard import CF_HTML_FORMAT_NAME, build_cf_html, extract_cf_html
from tabledown_windows.i18n import SUPPORTED_LANGUAGES, detect_system_language, t


MARKDOWN_BASIC = "| Name | Score |\n| --- | --- |\n| Alice | 95 |"
HTML_BASIC = "<table><tr><th>Name</th><th>Score</th></tr><tr><td>Alice</td><td>95</td></tr></table>"
# A table embedded in a document (paragraphs around it) — mirrors HTML_NOISE in
# scripts/run_test_matrix.py so both ports exercise the same document case.
HTML_DOCUMENT = (
    "<html><body><p>Before</p>"
    "<table><tr><th>제품</th><th>수량</th></tr><tr><td>키보드</td><td>3</td></tr></table>"
    "<p>After</p></body></html>"
)


class WindowsPortTests(unittest.TestCase):
    def test_cf_html_roundtrip_preserves_utf8_fragment(self):
        payload = build_cf_html("<table><tr><td>한글</td></tr></table>")

        html = extract_cf_html(payload)

        self.assertIn("<table>", html)
        self.assertIn("한글", html)

    def test_bare_excel_table_augments_text_and_keeps_html(self):
        # Invariant 3 (unified in 0.2.4): a bare Excel/Sheets table gains a
        # Markdown text slot but KEEPS CF_HTML (only rendered images are
        # dropped), so re-pasting into Excel/Word still yields a real table.
        result = converted_clipboard({"html": HTML_BASIC, "text": "Name\tScore\nAlice\t95"})

        self.assertEqual(result["text"], "\n| Name | Score |\n| --- | --- |\n| Alice | 95 |\n")
        self.assertIsNone(result.get("html"))
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])
        self.assertEqual(result["drop_formats"], WINDOWS_DROP_FORMATS)

    def test_html_table_with_mismatched_markdown_text_preserved(self):
        # Invariant 1: markdown text + html <table> is a real web/chat table. A
        # mismatched separator row must NOT trigger html->md (which would strip
        # html and break Excel paste); the clipboard is left untouched.
        result = converted_clipboard(
            {"html": HTML_BASIC, "text": "| # | A | B |\n| --- | --- |\n| 1 | x | y |"}
        )

        self.assertIsNone(result)

    def test_table_in_document_augments_text_and_keeps_html(self):
        # Invariant 4: a table inside a document augments the text slot with a
        # Markdown table while keeping the surrounding text, and preserves CF_HTML.
        result = converted_clipboard(
            {"html": HTML_DOCUMENT, "text": "Before\n제품 수량\n키보드 3\nAfter"}
        )

        self.assertIn("| 제품 | 수량 |", result["text"])
        self.assertIn("Before", result["text"])
        self.assertIn("After", result["text"])
        self.assertIsNone(result.get("html"))
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])

    def test_markdown_clipboard_adds_html_table(self):
        result = converted_clipboard({"text": MARKDOWN_BASIC})

        self.assertEqual(result["text"], MARKDOWN_BASIC)
        self.assertIn("<table><tr><th>Name</th><th>Score</th></tr>", result["html"])

    def test_generated_clipboard_is_skipped(self):
        self.assertIsNone(converted_clipboard({"generated": True, "text": MARKDOWN_BASIC}))

    def test_i18n_fallbacks(self):
        self.assertIn(detect_system_language(), SUPPORTED_LANGUAGES)
        self.assertEqual(t("menu.help", "ko"), "도움말")
        self.assertEqual(t("menu.help", "fr"), "Help")
        self.assertEqual(t("missing.key", "ko"), "missing.key")


if __name__ == "__main__":
    unittest.main()
