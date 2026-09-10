"""Cancellation contracts for macOS formula export, without native clipboard IO."""

from __future__ import annotations

import sys
import unittest
from unittest.mock import Mock, patch

if sys.platform == "darwin":
    from tablemark.app import TabledownApp
    from tablemark.clipboard import ClipboardChangedError
    from tablemark.excel_formula import ExcelFormulaError
    from tablemark.i18n import t
    from tests import test_excel_formula_macos as legacy_tests


@unittest.skipUnless(sys.platform == "darwin", "requires macOS app dependencies")
class FormulaCancellationActionTests(unittest.TestCase):
    def setUp(self):
        # Guard even accidental unpatched routes from touching user clipboard.
        self.pasteboards = self._patch("tablemark.clipboard.NSPasteboard")
        self.pasteboards.generalPasteboard.side_effect = AssertionError(
            "system clipboard access is forbidden in these tests"
        )
        self.app = legacy_tests.AppFormulaActionTests._app("ko")
        self.selection = legacy_tests.AppFormulaActionTests._notice_selection(
            partial=True,
        )
        self.generation = 23
        self.count = self._patch(
            "tablemark.app.clipboard_change_count",
            side_effect=lambda: self.generation,
        )
        self.schedule = self._patch(
            "tablemark.app.AppHelper.callAfter",
            side_effect=lambda callback, *args: callback(*args),
        )
        self.reader = self._patch(
            "tablemark.app.read_stable_selected_excel_formulas",
            return_value=self.selection,
        )
        self.serializer = self._patch(
            "tablemark.app.formula_selection_to_ai_xml",
            return_value="<표범위 />",
        )
        self.writer = self._patch("tablemark.app.write_text_only_clipboard")
        self.clipboard_reader = self._patch("tablemark.app.read_clipboard")
        self.automatic_writer = self._patch("tablemark.app.write_clipboard")
        self.table_writer = self._patch("tablemark.app.write_table_clipboard")
        self.logger = self._patch("tablemark.app.log")

    def tearDown(self):
        self.pasteboards.generalPasteboard.assert_not_called()
        self.automatic_writer.assert_not_called()
        self.table_writer.assert_not_called()
        # Neither failures nor success logging may include input values/names.
        diagnostics = repr(self.logger.call_args_list)
        self.assertNotIn("PrivateBook", diagnostics)
        self.assertNotIn("private pasted payload", diagnostics)

    def _patch(self, target, **kwargs):
        patcher = patch(target, **kwargs)
        result = patcher.start()
        self.addCleanup(patcher.stop)
        return result

    def _run(self):
        self.app.copy_selected_excel_formulas(None)

    def _assert_no_output(self):
        self.serializer.assert_not_called()
        self.writer.assert_not_called()
        self.app._flash_icon_success.assert_not_called()
        self.assertFalse(self.app._explicit_export_lock.locked())

    def _assert_changed_alert(self):
        self.app._safe_alert.assert_called_once_with(
            "Tabledown", t("formula.error.clipboard_changed", "ko"),
        )

    def test_callback_captures_original_generation_and_checks_current_stop_event(self):
        self._run()
        self.reader.assert_called_once()
        self.assertEqual(self.reader.call_args.args, ())
        self.assertEqual(set(self.reader.call_args.kwargs), {"check_cancelled"})
        check_cancelled = self.reader.call_args.kwargs["check_cancelled"]
        self.assertTrue(callable(check_cancelled))
        check_cancelled()

        self.generation = 24
        with self.assertRaises(ExcelFormulaError) as changed:
            check_cancelled()
        self.assertEqual(changed.exception.code, "clipboard_changed")

        self.generation = 23
        self.app._stop_watcher.set()
        with self.assertRaises(ExcelFormulaError) as stopped:
            check_cancelled()
        self.assertEqual(stopped.exception.code, "cancelled")

    def test_new_copy_before_queued_action_skips_excel_reader(self):
        queued = []
        self.schedule.side_effect = lambda callback, *args: queued.append(
            (callback, args)
        )
        self._run()
        self.assertEqual(len(queued), 1)
        self.generation = 24
        callback, args = queued.pop()
        callback(*args)

        self.reader.assert_not_called()
        self._assert_no_output()
        self._assert_changed_alert()
        self.assertIsNone(self.app._explicit_export_active)

    def test_stop_before_queued_action_skips_reader_and_notices(self):
        queued = []
        self.schedule.side_effect = lambda callback, *args: queued.append(
            (callback, args)
        )
        self._run()
        self.app._stop_watcher.set()
        callback, args = queued.pop()
        callback(*args)

        self.reader.assert_not_called()
        self._assert_no_output()
        self.app._safe_alert.assert_not_called()

    def test_new_copy_between_reader_batches_aborts_without_partial_success(self):
        events = []

        def read(*, check_cancelled):
            check_cancelled()
            events.append("first_batch")
            self.generation = 24
            check_cancelled()
            events.append("next_batch")
            return self.selection

        self.reader.side_effect = read
        self._run()

        self.assertEqual(events, ["first_batch"])
        self.reader.assert_called_once()
        self._assert_no_output()
        self._assert_changed_alert()
        self.assertIsNone(self.app._explicit_export_active)

    def test_stop_between_reader_batches_aborts_without_error_or_success_notice(self):
        events = []

        def read(*, check_cancelled):
            check_cancelled()
            events.append("first_batch")
            self.app._stop_watcher.set()
            check_cancelled()
            events.append("next_batch")
            return self.selection

        self.reader.side_effect = read
        self._run()

        self.assertEqual(events, ["first_batch"])
        self.reader.assert_called_once()
        self._assert_no_output()
        self.app._safe_alert.assert_not_called()

    def test_new_copy_at_reader_return_is_checked_before_serialization(self):
        def read(*, check_cancelled):
            check_cancelled()
            self.generation = 24
            return self.selection

        self.reader.side_effect = read
        self._run()

        self._assert_no_output()
        self._assert_changed_alert()
        self.assertIsNone(self.app._explicit_export_active)

    def test_stop_at_reader_return_is_checked_before_serialization(self):
        def read(*, check_cancelled):
            check_cancelled()
            self.app._stop_watcher.set()
            return self.selection

        self.reader.side_effect = read
        self._run()

        self._assert_no_output()
        self.app._safe_alert.assert_not_called()

    def test_copy_after_last_checkpoint_keeps_writer_generation_guard_and_no_notice(self):
        def write(*args, **kwargs):
            # The external copy occurs after all reader/application checkpoints.
            # The production writer owns this last generation check.
            self.generation = 24
            raise ClipboardChangedError("private pasted payload")

        self.writer.side_effect = write
        self._run()

        self.serializer.assert_called_once_with(self.selection)
        self.writer.assert_called_once_with(
            "<표범위 />", mark_generated=True, expected_change_count=23,
        )
        self._assert_changed_alert()
        self.app._flash_icon_success.assert_not_called()
        self.assertIsNone(self.app._explicit_export_active)
        self.assertFalse(self.app._explicit_export_lock.locked())

    def test_cancel_releases_gate_and_manual_retry_captures_new_generation(self):
        callbacks = []

        def read(*, check_cancelled):
            callbacks.append(check_cancelled)
            check_cancelled()
            if len(callbacks) == 1:
                self.generation = 24
                check_cancelled()
            return self.selection

        self.reader.side_effect = read
        self._run()
        self._assert_no_output()
        self._assert_changed_alert()
        self.assertIsNone(self.app._explicit_export_active)
        for item in (
            self.app.copy_markdown_item,
            self.app.copy_xml_item,
            self.app.copy_excel_formulas_item,
        ):
            item._menuitem.setEnabled_.assert_called_with(True)

        self.app._safe_alert.reset_mock()
        self._run()
        self.assertEqual(self.reader.call_count, 2)
        self.assertEqual(len(callbacks), 2)
        self.assertIsNot(callbacks[0], callbacks[1])
        self.serializer.assert_called_once_with(self.selection)
        self.writer.assert_called_once_with(
            "<표범위 />", mark_generated=True, expected_change_count=24,
        )
        self.app._flash_icon_success.assert_called_once_with()
        self.app._safe_alert.assert_called_once_with(
            t("formula.copy_notice_title", "ko"),
            t("formula.copy_notice.partial_references", "ko"),
        )
        self.assertIsNone(self.app._explicit_export_active)
        self.assertFalse(self.app._explicit_export_lock.locked())

    def test_cancel_leaves_new_generation_for_next_watcher_tick(self):
        self.app.enabled = True
        self.app._last_change_count = 23
        self.app._augment_clipboard = TabledownApp._augment_clipboard.__get__(
            self.app
        )
        self.app._converted_clipboard = Mock(return_value=None)
        deferred_results = []

        def read(*, check_cancelled):
            self.generation = 24
            deferred_results.append(self.app._augment_clipboard())
            check_cancelled()
            return self.selection

        self.reader.side_effect = read
        self._run()
        self.assertEqual(deferred_results, [False])
        self.assertEqual(self.app._last_change_count, 23)
        self.clipboard_reader.assert_not_called()
        self._assert_no_output()
        self._assert_changed_alert()

        new_copy = {"text": "a newly copied non-table value", "html": ""}
        self.clipboard_reader.return_value = new_copy
        ticks = []

        def tick(_seconds):
            ticks.append(None)
            if len(ticks) == 2:
                self.app.enabled = False
                self.app._stop_watcher.set()

        with patch("tablemark.app.time.sleep", side_effect=tick):
            TabledownApp._watch_clipboard(self.app)

        self.assertEqual(len(ticks), 2)
        self.clipboard_reader.assert_called_once_with()
        self.app._converted_clipboard.assert_called_once_with(new_copy)
        self.assertEqual(self.app._last_change_count, 24)
        self.writer.assert_not_called()


if __name__ == "__main__":
    unittest.main()
