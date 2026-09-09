"""Check compact AI export through the existing macOS action lifecycle."""

import sys
import unittest
from unittest.mock import Mock, patch

if sys.platform == "darwin":
    from tablemark import clipboard
    from tablemark.app import TabledownApp
    from tablemark.converter.formula_export import FormulaXmlTooLargeError
    from tablemark.excel_formula import ExcelFormulaError, SELECTION_CHANGED
    from tablemark.i18n import t
    from tests import test_excel_formula_macos as legacy_tests


@unittest.skipUnless(sys.platform == "darwin", "requires macOS app dependencies")
class AiFormulaActionTests(unittest.TestCase):
    def _app(self):
        app = legacy_tests.AppFormulaActionTests._app("ko")
        app.copy_selected_excel_formulas_for_ai = (
            TabledownApp.copy_selected_excel_formulas_for_ai.__get__(app)
        )
        return app

    @staticmethod
    def _run(app):
        with (
            patch("tablemark.app.clipboard_change_count", return_value=23),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=lambda callback, *args: callback(*args),
            ),
        ):
            app.copy_selected_excel_formulas_for_ai(None)

    def test_ai_action_selects_new_serializer_and_reads_once(self):
        app = self._app()
        selection = legacy_tests.AppFormulaActionTests._notice_selection(partial=False)
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection) as reader,
            patch("tablemark.app.formula_selection_to_ai_xml", return_value="<표범위 />") as compact,
            patch("tablemark.app.formula_selection_to_xml") as original,
            patch("tablemark.app.write_text_only_clipboard") as writer,
        ):
            self._run(app)
        reader.assert_called_once_with()
        compact.assert_called_once_with(selection)
        original.assert_not_called()
        writer.assert_called_once_with("<표범위 />", mark_generated=True, expected_change_count=23)
        app._flash_icon_success.assert_called_once_with()
        app._safe_alert.assert_not_called()
        self.assertIsNone(app._explicit_export_active)
        self.assertFalse(app._explicit_export_lock.locked())

    def test_existing_action_keeps_original_serializer(self):
        app = self._app()
        selection = legacy_tests.AppFormulaActionTests._notice_selection(partial=False)
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection),
            patch("tablemark.app.formula_selection_to_ai_xml") as compact,
            patch("tablemark.app.formula_selection_to_xml", return_value="<表 />") as original,
            patch("tablemark.app.write_text_only_clipboard"),
        ):
            legacy_tests.AppFormulaActionTests._run(app)
        compact.assert_not_called()
        original.assert_called_once_with(selection)

    def test_all_export_actions_share_gate_and_busy_menu(self):
        app = self._app()
        queued = []
        with (
            patch("tablemark.app.clipboard_change_count", return_value=23),
            patch("tablemark.app.AppHelper.callAfter", side_effect=lambda cb, *args: queued.append((cb, args))),
            patch("tablemark.app.read_stable_selected_excel_formulas") as reader,
        ):
            app.copy_selected_excel_formulas_for_ai(None)
            self.assertEqual(app.copy_ai_formulas_item.title, t("menu.copy_ai_formulas_busy", "ko"))
            for item in (app.copy_ai_formulas_item, app.copy_excel_formulas_item, app.copy_xml_item):
                item._menuitem.setEnabled_.assert_called_with(False)
            app.copy_selected_excel_formulas_for_ai(None)
            app.copy_selected_excel_formulas(None)
            TabledownApp.copy_as_xml(app, None)
            self.assertEqual(len(queued), 1)
            self.assertEqual(app._safe_alert.call_count, 3)
            reader.assert_not_called()
            # Ending before the queued callback must release the shared gate.
            app._stop_watcher.set()
            callback, args = queued[0]
            callback(*args)
            reader.assert_not_called()
        self.assertFalse(app._explicit_export_lock.locked())
        app._flash_icon_success.assert_not_called()

    def test_stop_during_read_does_not_write_or_report_success(self):
        app = self._app()
        selection = legacy_tests.AppFormulaActionTests._notice_selection()
        def stop_after_read():
            app._stop_watcher.set()
            return selection
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", side_effect=stop_after_read),
            patch("tablemark.app.formula_selection_to_ai_xml", return_value="<표범위 />"),
            patch("tablemark.app.write_text_only_clipboard") as writer,
        ):
            self._run(app)
        writer.assert_not_called()
        app._safe_alert.assert_not_called()
        app._flash_icon_success.assert_not_called()
        self.assertFalse(app._explicit_export_lock.locked())

    def test_failures_do_not_report_success_and_release_gate(self):
        selection = legacy_tests.AppFormulaActionTests._notice_selection(partial=False)
        cases = (
            ("reader", ExcelFormulaError(SELECTION_CHANGED), "selection_changed"),
            ("serializer", ValueError("private cell text"), "execution_failed"),
            ("serializer", FormulaXmlTooLargeError("private cell text"), "output_too_large"),
            ("writer", clipboard.ClipboardChangedError("private cell text"), "clipboard_changed"),
            ("writer", clipboard.ClipboardWriteError("private cell text"), "clipboard_write_failed"),
        )
        for target, error, code in cases:
            with self.subTest(target=target, code=code):
                app = self._app()
                with (
                    patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection) as reader,
                    patch("tablemark.app.formula_selection_to_ai_xml", return_value="<표범위 />") as serializer,
                    patch("tablemark.app.write_text_only_clipboard") as writer,
                    patch("tablemark.app.log") as logger,
                ):
                    {"reader": reader, "serializer": serializer, "writer": writer}[target].side_effect = error
                    self._run(app)
                if target != "writer":
                    writer.assert_not_called()
                app._flash_icon_success.assert_not_called()
                app._safe_alert.assert_called_once_with("Tabledown", t(f"formula.error.{code}", "ko"))
                self.assertNotIn("private cell text", repr(logger.call_args_list))
                self.assertFalse(app._explicit_export_lock.locked())
                self.assertIsNone(app._explicit_export_active)
                app.copy_ai_formulas_item._menuitem.setEnabled_.assert_called_with(True)

    def test_partial_notice_only_after_verified_write_and_manual_retry(self):
        app = self._app()
        selection = legacy_tests.AppFormulaActionTests._notice_selection()
        events = []
        app._safe_alert.side_effect = lambda *_: events.append("notice")
        def write(*args, **kwargs):
            events.append("write")
            if len(events) == 1:
                raise clipboard.ClipboardWriteError("content-free")
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection),
            patch("tablemark.app.formula_selection_to_ai_xml", return_value="<표범위 />"),
            patch("tablemark.app.write_text_only_clipboard", side_effect=write) as writer,
        ):
            self._run(app)
            self.assertFalse(app._explicit_export_lock.locked())
            app._flash_icon_success.assert_not_called()
            self._run(app)
        self.assertEqual(events, ["write", "notice", "write", "notice"])
        self.assertEqual(writer.call_count, 2)
        app._flash_icon_success.assert_called_once_with()
        self.assertEqual(app._safe_alert.call_args.args[1], t("formula.copy_notice.partial_references", "ko"))

    def test_schedule_failure_and_language_change_restore_ai_label(self):
        app = self._app()
        with (
            patch("tablemark.app.clipboard_change_count", return_value=23),
            patch("tablemark.app.AppHelper.callAfter", side_effect=RuntimeError("queue unavailable")),
            patch("tablemark.app.write_text_only_clipboard") as writer,
        ):
            app.copy_selected_excel_formulas_for_ai(None)
        writer.assert_not_called()
        self.assertFalse(app._explicit_export_lock.locked())
        app.lang = "en"
        app._update_explicit_export_menu()
        self.assertEqual(app.copy_ai_formulas_item.title, "Copy compact XML for AI")
        app.copy_ai_formulas_item._menuitem.setToolTip_.assert_called_with(t("menu.copy_ai_formulas_tooltip", "en"))


if __name__ == "__main__":
    unittest.main()
