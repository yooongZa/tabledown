from __future__ import annotations

import contextlib
import sys
import unittest
from pathlib import Path
from unittest import mock


WINDOWS_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = WINDOWS_ROOT.parent
for path in (PROJECT_ROOT, WINDOWS_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from tabledown_windows import single_instance, startup_task
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


def _excel_style_cf_html(interior: str) -> bytes:
    """Build CF_HTML the way Excel/Sheets do: the StartFragment/EndFragment
    markers sit *inside* the table, so the fragment is the table interior
    (``<col>``/``<tr>``/``<td>``) with the ``<table>`` tag left outside it.
    """
    pre = "<html>\r\n<body>\r\n<table border=0>"  # <table> precedes the fragment
    post = "</table>\r\n</body>\r\n</html>"
    html = pre + interior + post

    header_tmpl = (
        "Version:1.0\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\n"
        "StartFragment:0000000000\r\nEndFragment:0000000000\r\n"
    )
    base = len(header_tmpl.encode("utf-8"))
    start_html = base
    start_fragment = base + len(pre.encode("utf-8"))  # points AFTER <table>
    end_fragment = start_fragment + len(interior.encode("utf-8"))
    end_html = base + len(html.encode("utf-8"))
    header = (
        f"Version:1.0\r\nStartHTML:{start_html:010d}\r\nEndHTML:{end_html:010d}\r\n"
        f"StartFragment:{start_fragment:010d}\r\nEndFragment:{end_fragment:010d}\r\n"
    )
    return header.encode("utf-8") + html.encode("utf-8")


class WindowsPortTests(unittest.TestCase):
    def test_excel_cf_html_fragment_without_table_tag_is_wrapped(self):
        # Excel's fragment omits the <table> tag (markers are inside it). Without
        # wrapping, table detection fails and an Excel copy never converts.
        payload = _excel_style_cf_html(
            "<col width=80><tr><td>동해물과</td><td>백두산이</td></tr>"
            "<tr><td>마르고</td><td>닳도록</td></tr>"
        )

        html = extract_cf_html(payload)

        self.assertIn("<table", html.lower())
        self.assertIn("동해물과", html)

    def test_real_excel_table_converts_to_markdown(self):
        # The end-to-end regression for the reported "Excel tables don't convert"
        # bug: a realistic Excel CF_HTML payload + its TSV text must yield a
        # Markdown table in the text slot (html kept, only rendered images dropped).
        payload = _excel_style_cf_html(
            "<col width=80 span=2>"
            "<tr><td>동해물과</td><td>백두산이</td></tr>"
            "<tr><td>마르고</td><td>닳도록</td></tr>"
        )
        result = converted_clipboard(
            {"html": extract_cf_html(payload), "text": "동해물과\t백두산이\r\n마르고\t닳도록"}
        )

        self.assertIsNotNone(result)
        self.assertIn("| 동해물과 | 백두산이 |", result["text"])
        self.assertIn("| 마르고 | 닳도록 |", result["text"])
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])

    def test_cf_html_roundtrip_preserves_utf8_fragment(self):
        payload = build_cf_html("<table><tr><td>한글</td></tr></table>")

        html = extract_cf_html(payload)

        self.assertIn("<table>", html)
        self.assertIn("한글", html)

    def test_decode_html_prefers_cp1252_over_utf16_mojibake(self):
        # An even-length, invalid-UTF-8 cp1252 fragment must decode as cp1252,
        # not be mis-read as UTF-16 (which "succeeds" and merges latin byte
        # pairs into CJK glyphs, then fails table detection — a conversion macOS
        # performs because it reads ready-made HTML and never mis-decodes).
        from tabledown_windows.html_clipboard import _decode_html

        raw = "café".encode("cp1252")  # b'caf\xe9' — invalid utf-8, even length
        self.assertEqual(_decode_html(raw), "café")

    def test_decode_html_honors_utf16_bom(self):
        # A genuine UTF-16 fragment (BOM present) is still decoded correctly.
        from tabledown_windows.html_clipboard import _decode_html

        self.assertEqual(_decode_html("한글표".encode("utf-16")), "한글표")

    def test_ensure_table_wrapper_ignores_non_row_tags(self):
        # <track>/<trail>/<tr-foo> share the "<tr" prefix but are not rows; they
        # must not be wrapped into a bogus <table>. A real <tr> still is.
        from tabledown_windows.html_clipboard import _ensure_table_wrapper

        track = "<video><track kind='subtitles'></video>"
        self.assertEqual(_ensure_table_wrapper(track), track)
        self.assertIn("<table>", _ensure_table_wrapper("<tr><td>x</td></tr>"))

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

    def test_login_item_translations_exist(self):
        # The menu label and both "couldn't enable" hints must resolve in each
        # language — not fall back to the bare key.
        self.assertEqual(t("menu.login_item", "ko"), "로그인 시 자동 실행")
        self.assertEqual(t("menu.login_item", "en"), "Open at Login")
        for lang in SUPPORTED_LANGUAGES:
            for key in ("login_item.blocked_by_user", "login_item.blocked_by_policy"):
                self.assertNotEqual(t(key, lang), key)


class _FakeWinrtTask:
    """Stand-in for a WinRT StartupTask object.

    ``state`` returns a real ``StartupTaskState`` member so the production
    ``_state_name`` mapping is exercised. ``request_enable_async`` returns a
    coroutine (awaitable, like the real ``IAsyncOperation``); ``disable`` is
    synchronous, matching the real API.
    """

    def __init__(self, state, becomes=None):
        self._state = state
        self._becomes = becomes
        self.enable_requested = False
        self.disabled = False

    @property
    def state(self):
        return self._state

    def request_enable_async(self):
        async def _op():
            self.enable_requested = True
            if self._becomes is not None:
                self._state = self._becomes
        return _op()

    def disable(self) -> None:
        self.disabled = True
        if self._becomes is not None:
            self._state = self._becomes


class _FakeStartupTaskApi:
    """Stand-in for the winsdk ``StartupTask`` class (its ``get_async``)."""

    def __init__(self, task: _FakeWinrtTask):
        self._task = task

    def get_async(self, _task_id):
        async def _op():
            return self._task
        return _op()


@contextlib.contextmanager
def _patched_api(task: _FakeWinrtTask):
    # Swap winsdk's StartupTask for a fake; the real _run/asyncio/_state_name
    # orchestration runs against it on the production code path.
    with mock.patch.object(startup_task, "StartupTask", _FakeStartupTaskApi(task)):
        yield


class StartupTaskDegradeTests(unittest.TestCase):
    def test_degrades_without_winrt(self):
        # No winsdk / no package identity: every entry point is safe and falsey,
        # so the menu just omits the toggle (mirrors macOS login_item).
        with mock.patch.object(startup_task, "StartupTask", None):
            self.assertIsNone(startup_task.current_state())
            self.assertFalse(startup_task.is_supported())
            self.assertFalse(startup_task.is_enabled())
            self.assertEqual(startup_task.set_enabled(True), "unavailable")


@unittest.skipUnless(
    startup_task.StartupTaskState is not None, "winsdk StartupTaskState required"
)
class StartupTaskTests(unittest.TestCase):
    def _state(self, name):
        return getattr(startup_task.StartupTaskState, name)

    def test_current_state_and_is_enabled(self):
        with _patched_api(_FakeWinrtTask(self._state("ENABLED"))):
            self.assertEqual(startup_task.current_state(), "enabled")
            self.assertTrue(startup_task.is_supported())
            self.assertTrue(startup_task.is_enabled())

    def test_enable_requests_and_reports_enabled(self):
        fake = _FakeWinrtTask(self._state("DISABLED"), becomes=self._state("ENABLED"))
        with _patched_api(fake):
            self.assertEqual(startup_task.set_enabled(True), "enabled")
        self.assertTrue(fake.enable_requested)

    def test_disable_calls_disable(self):
        fake = _FakeWinrtTask(self._state("ENABLED"), becomes=self._state("DISABLED"))
        with _patched_api(fake):
            self.assertEqual(startup_task.set_enabled(False), "disabled")
        self.assertTrue(fake.disabled)

    def test_blocked_by_user_is_not_an_enabled_state(self):
        # Windows can keep an enable request DISABLED_BY_USER; the read-back name
        # must surface that and never count as enabled.
        fake = _FakeWinrtTask(self._state("DISABLED_BY_USER"))  # stays put
        with _patched_api(fake):
            status = startup_task.set_enabled(True)
        self.assertTrue(fake.enable_requested)
        self.assertEqual(status, "disabled_by_user")
        self.assertNotIn(status, startup_task.ENABLED_STATES)


@unittest.skipUnless(sys.platform.startswith("win"), "tray app is Windows-only")
class LoginMenuTests(unittest.TestCase):
    def _make_app(self, *, supported: bool, enabled: bool = False):
        # app.__init__ seeds login_supported/login_enabled from one
        # current_state() read; None means "unsupported, hide the toggle".
        state = ("enabled" if enabled else "disabled") if supported else None
        with mock.patch.object(startup_task, "current_state", return_value=state):
            from tabledown_windows.app import TabledownWindowsApp

            return TabledownWindowsApp()

    def _labels(self, app):
        return [getattr(item, "text", "") for item in app.icon.menu]

    def test_menu_includes_login_item_when_supported(self):
        app = self._make_app(supported=True)
        self.assertIn(t("menu.login_item", app.lang), self._labels(app))

    def test_menu_omits_login_item_when_unsupported(self):
        app = self._make_app(supported=False)
        self.assertNotIn(t("menu.login_item", app.lang), self._labels(app))

    def test_toggle_blocked_by_user_stays_unchecked_and_warns(self):
        app = self._make_app(supported=True, enabled=False)
        shown = []
        with mock.patch.object(startup_task, "set_enabled", return_value="disabled_by_user"), \
             mock.patch.object(app, "_refresh_menu"), \
             mock.patch.object(
                 app, "_show_message_box_async", side_effect=lambda msg, _title: shown.append(msg)
             ):
            app.toggle_login_item(None, None)
        self.assertFalse(app.login_enabled)  # checkmark stays truthful
        self.assertEqual(len(shown), 1)      # user is told why

    def test_toggle_enable_checks_and_is_silent(self):
        app = self._make_app(supported=True, enabled=False)
        shown = []
        with mock.patch.object(startup_task, "set_enabled", return_value="enabled"), \
             mock.patch.object(app, "_refresh_menu"), \
             mock.patch.object(
                 app, "_show_message_box_async", side_effect=lambda msg, _title: shown.append(msg)
             ):
            app.toggle_login_item(None, None)
        self.assertTrue(app.login_enabled)
        self.assertEqual(shown, [])  # success is silent — the checkmark says it


class SingleInstanceTests(unittest.TestCase):
    """The named-mutex guard that stops a second tray (and second watcher)."""

    def setUp(self):
        # acquire_single_instance stashes the live handle in a module global;
        # clear it so one test's claim can't leak into the next.
        single_instance._held_handle = None
        self.addCleanup(setattr, single_instance, "_held_handle", None)

    def test_first_instance_acquires_and_holds_handle(self):
        # Fresh mutex (not pre-existing) -> this is the owner; the handle is kept
        # alive for the process lifetime so the mutex outlives the call.
        with mock.patch.object(single_instance, "_create_mutex", return_value=(4321, False)):
            self.assertTrue(single_instance.acquire_single_instance())
        self.assertEqual(single_instance._held_handle, 4321)

    def test_second_instance_is_blocked(self):
        # CreateMutexW reported ERROR_ALREADY_EXISTS -> a sibling owns it; bail
        # out and do not retain a handle.
        with mock.patch.object(single_instance, "_create_mutex", return_value=(4321, True)):
            self.assertFalse(single_instance.acquire_single_instance())
        self.assertIsNone(single_instance._held_handle)

    def test_missing_kernel_degrades_to_allowed(self):
        # Non-Windows host (no kernel32): _create_mutex yields (None, False), so
        # the app is allowed to start rather than refusing over a missing guard.
        with mock.patch.object(single_instance, "_create_mutex", return_value=(None, False)):
            self.assertTrue(single_instance.acquire_single_instance())


@unittest.skipUnless(sys.platform.startswith("win"), "tray app is Windows-only")
class SingleInstanceMainTests(unittest.TestCase):
    def test_main_skips_tray_when_already_running(self):
        # The guard must short-circuit main() before the tray app is built, so a
        # second launch adds neither an icon nor a clipboard watcher.
        from tabledown_windows import app

        with mock.patch.object(app.single_instance, "acquire_single_instance", return_value=False), \
             mock.patch.object(app, "TabledownWindowsApp") as fake_app:
            app.main()
        fake_app.assert_not_called()

    def test_main_builds_tray_when_first(self):
        from tabledown_windows import app

        with mock.patch.object(app.single_instance, "acquire_single_instance", return_value=True), \
             mock.patch.object(app, "TabledownWindowsApp") as fake_app:
            app.main()
        fake_app.assert_called_once_with()
        fake_app.return_value.run.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
