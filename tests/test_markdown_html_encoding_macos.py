"""Read generated clipboard HTML through macOS's actual HTML importer."""

import sys
import unittest
from unittest.mock import Mock, patch

if sys.platform == "darwin":
    from AppKit import (
        NSAttributedString,
        NSDocumentTypeDocumentAttribute,
        NSHTMLTextDocumentType,
        NSPasteboard,
        NSPasteboardTypeString,
    )

    from tablemark.app import TabledownApp
    from tablemark.clipboard import (
        GENERATED_MARKER_TYPES,
        HTML_TYPES,
        STRING_TYPES,
        ClipboardChangedError,
        ClipboardWriteError,
        write_clipboard,
        write_table_clipboard,
    )


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

    def test_explicit_table_preserves_unicode_merge_and_breaks_in_native_import(self):
        markdown = (
            "| 항목 | 값 |\n| --- | --- |\n"
            "| 한글 🍎 | 첫째<br>둘째 |\n|  | A&B <태그> |"
        )
        html = (
            '<meta charset="utf-8"><table>'
            "<tr><th>항목</th><th>값</th></tr>"
            '<tr><td rowspan="2">한글 🍎</td><td>첫째<br>둘째</td></tr>'
            "<tr><td>A&amp;B &lt;태그&gt;</td></tr></table>"
        )
        stale_types = {
            "com.microsoft.Excel.biff12", "public.png", "public.rtf"
        }
        pasteboard = NSPasteboard.pasteboardWithUniqueName()
        try:
            pasteboard.declareTypes_owner_(sorted(stale_types), None)
            for pb_type in stale_types:
                pasteboard.setString_forType_("old synthetic copy", pb_type)
            generation = int(pasteboard.changeCount())
            with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                pasteboards.generalPasteboard.return_value = pasteboard
                write_table_clipboard(
                    markdown, html, expected_change_count=generation
                )

            self.assertEqual(
                {str(pb_type) for pb_type in pasteboard.types()},
                STRING_TYPES | HTML_TYPES | set(GENERATED_MARKER_TYPES),
            )
            for pb_type in stale_types:
                self.assertIsNone(pasteboard.dataForType_(pb_type))
            for pb_type in STRING_TYPES:
                self.assertEqual(str(pasteboard.stringForType_(pb_type)), markdown)
            for pb_type, expected in GENERATED_MARKER_TYPES.items():
                self.assertEqual(str(pasteboard.stringForType_(pb_type)), expected)
            for html_type in HTML_TYPES:
                with self.subTest(html_type=html_type):
                    self.assertEqual(str(pasteboard.stringForType_(html_type)), html)
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
                    self.assertEqual(
                        str(value.string()).splitlines(),
                        ["항목", "값", "한글 🍎", "첫째", "둘째", "A&B <태그>"],
                    )
        finally:
            pasteboard.releaseGlobally()

    def test_explicit_table_can_omit_generated_markers(self):
        pasteboard = NSPasteboard.pasteboardWithUniqueName()
        try:
            with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                pasteboards.generalPasteboard.return_value = pasteboard
                write_table_clipboard("한글", "<table></table>", mark_generated=False)
            self.assertEqual(
                {str(pb_type) for pb_type in pasteboard.types()},
                STRING_TYPES | HTML_TYPES,
            )
        finally:
            pasteboard.releaseGlobally()

    def test_explicit_table_keeps_newer_copy_when_generation_changed(self):
        pasteboard = NSPasteboard.pasteboardWithUniqueName()
        try:
            original_generation = int(pasteboard.changeCount())
            pasteboard.clearContents()
            pasteboard.declareTypes_owner_([NSPasteboardTypeString], None)
            pasteboard.setString_forType_("new synthetic copy", NSPasteboardTypeString)
            newer_generation = int(pasteboard.changeCount())
            with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                pasteboards.generalPasteboard.return_value = pasteboard
                with self.assertRaisesRegex(ClipboardChangedError, "^clipboard_changed$"):
                    write_table_clipboard(
                        "older export", "<table></table>",
                        expected_change_count=original_generation,
                    )
            self.assertEqual(int(pasteboard.changeCount()), newer_generation)
            self.assertEqual(
                str(pasteboard.stringForType_(NSPasteboardTypeString)),
                "new synthetic copy",
            )
        finally:
            pasteboard.releaseGlobally()

    @staticmethod
    def _mock_pasteboard():
        data = {}
        pasteboard = Mock()
        pasteboard.changeCount.return_value = 17
        pasteboard.clearContents.side_effect = data.clear

        def write(value, pb_type):
            data[pb_type] = value
            return True

        pasteboard.setString_forType_.side_effect = write
        pasteboard.stringForType_.side_effect = data.get
        return pasteboard, data

    def test_explicit_table_rechecks_generation_immediately_before_clear(self):
        pasteboard, _ = self._mock_pasteboard()
        pasteboard.changeCount.side_effect = [17, 18]
        with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
            pasteboards.generalPasteboard.return_value = pasteboard
            with self.assertRaisesRegex(ClipboardChangedError, "^clipboard_changed$"):
                write_table_clipboard("new", "<table></table>", expected_change_count=17)
        pasteboard.clearContents.assert_not_called()
        pasteboard.declareTypes_owner_.assert_not_called()
        pasteboard.setString_forType_.assert_not_called()

    def test_explicit_table_reports_false_write_for_every_format_without_retry(self):
        for failed_type in STRING_TYPES | HTML_TYPES | set(GENERATED_MARKER_TYPES):
            with self.subTest(failed_type=failed_type):
                pasteboard, data = self._mock_pasteboard()

                def write(value, pb_type):
                    if pb_type == failed_type:
                        return False
                    data[pb_type] = value
                    return True

                pasteboard.setString_forType_.side_effect = write
                with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                    pasteboards.generalPasteboard.return_value = pasteboard
                    with self.assertRaisesRegex(ClipboardWriteError, "^clipboard_write_failed$"):
                        write_table_clipboard("한글", "<table></table>")
                pasteboard.clearContents.assert_called_once()
                pasteboard.stringForType_.assert_not_called()
                self.assertEqual(
                    sum(call.args[1] == failed_type for call in pasteboard.setString_forType_.call_args_list),
                    1,
                )

    def test_explicit_table_reports_missing_or_different_readback_for_every_format(self):
        for failed_type in STRING_TYPES | HTML_TYPES | set(GENERATED_MARKER_TYPES):
            for incorrect_value in (None, "different synthetic content"):
                with self.subTest(failed_type=failed_type, missing=incorrect_value is None):
                    pasteboard, data = self._mock_pasteboard()
                    pasteboard.stringForType_.side_effect = (
                        lambda pb_type: incorrect_value if pb_type == failed_type else data.get(pb_type)
                    )
                    with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                        pasteboards.generalPasteboard.return_value = pasteboard
                        with self.assertRaisesRegex(ClipboardWriteError, "^clipboard_write_failed$"):
                            write_table_clipboard("한글", "<table></table>")
                    pasteboard.clearContents.assert_called_once()

    def test_explicit_table_hides_system_exception_contents_and_does_not_retry(self):
        for method in (
            "changeCount", "clearContents", "declareTypes_owner_",
            "setString_forType_", "stringForType_",
        ):
            with self.subTest(method=method):
                pasteboard, _ = self._mock_pasteboard()
                getattr(pasteboard, method).side_effect = RuntimeError("private clipboard contents")
                with patch("tablemark.clipboard.NSPasteboard") as pasteboards:
                    pasteboards.generalPasteboard.return_value = pasteboard
                    with self.assertRaisesRegex(ClipboardWriteError, "^clipboard_write_failed$"):
                        write_table_clipboard("한글", "<table></table>", expected_change_count=17)
                self.assertLessEqual(pasteboard.clearContents.call_count, 1)


if __name__ == "__main__":
    unittest.main()
