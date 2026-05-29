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


class WindowsPortTests(unittest.TestCase):
    def test_cf_html_roundtrip_preserves_utf8_fragment(self):
        payload = build_cf_html("<table><tr><td>한글</td></tr></table>")

        html = extract_cf_html(payload)

        self.assertIn("<table>", html)
        self.assertIn("한글", html)

    def test_html_clipboard_converts_to_markdown_and_drops_windows_html(self):
        result = converted_clipboard({"html": HTML_BASIC, "text": "Name\tScore\nAlice\t95"})

        self.assertEqual(result["text"], "\n| Name | Score |\n| --- | --- |\n| Alice | 95 |\n")
        self.assertIn(CF_HTML_FORMAT_NAME, result["drop_formats"])
        self.assertEqual(result["drop_formats"], WINDOWS_DROP_FORMATS)

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
