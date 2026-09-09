"""Read generated clipboard HTML through macOS's actual HTML importer."""

import sys
import unittest
from unittest.mock import patch

if sys.platform == "darwin":
    from AppKit import (
        NSAttributedString,
        NSDocumentTypeDocumentAttribute,
        NSHTMLTextDocumentType,
        NSPasteboard,
        NSPasteboardTypeString,
    )

    from tablemark.app import TabledownApp
    from tablemark.clipboard import HTML_TYPES, write_clipboard


@unittest.skipUnless(sys.platform == "darwin", "requires macOS HTML importer")
class MarkdownHtmlEncodingTests(unittest.TestCase):
    def test_clipboard_html_preserves_unicode_in_native_rich_text_import(self):
        markdown = (
            "| 항목 | 값 |\n| --- | --- |\n"
            "| 한글 확인 | 사과 🍎 |\n"
            '| 특수문자 | A&B <태그> "인용" |\n'
            "| English | 123 |"
        )
        expected = [
            "항목", "값", "한글 확인", "사과 🍎",
            "특수문자", 'A&B <태그> "인용"', "English", "123",
        ]
        update = TabledownApp._converted_clipboard(None, {"text": markdown})
        self.assertEqual(update["text"], markdown)

        # Exercise the production writer without changing the user's clipboard.
        pasteboard = NSPasteboard.pasteboardWithUniqueName()
        try:
            with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                pasteboards.generalPasteboard.return_value = pasteboard
                write_clipboard(**update, mark_generated=True)
            self.assertEqual(
                str(pasteboard.stringForType_(NSPasteboardTypeString)), markdown
            )
            for html_type in sorted(HTML_TYPES):
                with self.subTest(html_type=html_type):
                    # A stringForType roundtrip cannot detect a rich-text reader
                    # guessing the wrong encoding. Let AppKit decode raw data
                    # without supplying an out-of-band character encoding.
                    value, _, error = (
                        NSAttributedString.alloc()
                        .initWithData_options_documentAttributes_error_(
                            pasteboard.dataForType_(html_type),
                            {NSDocumentTypeDocumentAttribute: NSHTMLTextDocumentType},
                            None,
                            None,
                        )
                    )
                    self.assertIsNone(error)
                    self.assertIsNotNone(value)
                    self.assertEqual(str(value.string()).splitlines(), expected)
        finally:
            pasteboard.releaseGlobally()


if __name__ == "__main__":
    unittest.main()
