"""Windows tray app for Tabledown."""
from __future__ import annotations

import ctypes
from pathlib import Path
import sys
import threading
import time

from PIL import Image, ImageChops, ImageDraw
import pystray

from tablemark.converter.formula_export import formula_selection_to_xml

from . import diagnostics, single_instance, startup_task
from .conversion import converted_clipboard
from .excel_formula import read_selected_excel_formulas
from .hotkey import MOD_ALT, MOD_CONTROL, MOD_NOREPEAT, VK_E, VK_T, GlobalHotkey
from .i18n import SUPPORTED_LANGUAGES, resolve_language, save_preferred_language, t
from .logger import log
from .settings import load_settings, save_setting
from .win_clipboard import (
    clipboard_change_count,
    read_clipboard,
    write_clipboard,
    write_text_only_clipboard,
)

WELCOME_SHOWN_KEY = "welcome_shown"
# Persisted "빈칸을 자동 채우기" preference (0.5.0, parity with macOS fill_blanks).
# Off by default so a merge stays blank unless the user opts in.
FILL_BLANKS_KEY = "fill_blanks"

# MessageBoxW flags: MB_ICONINFORMATION.
_MB_ICONINFORMATION = 0x00000040
# Pull the dialog to the foreground and keep it on top. A tray callback runs
# with no foreground rights, so without these the box can open *behind* the
# active window — the user clicks Help again, stacking boxes that feel "stuck".
_MB_SETFOREGROUND = 0x00010000
_MB_TOPMOST = 0x00040000
# GetSystemMetrics index for the small-icon width the tray actually uses
# (DPI-scaled by Windows: 16 at 96dpi, larger on HiDPI).
_SM_CXSMICON = 49


class TabledownWindowsApp:
    """Windows notification area app."""

    def __init__(self):
        log("windows app starting")
        self.enabled = True
        self.lang = resolve_language()
        # Persisted preference (parity with macOS): drives the Markdown path's
        # merged-header fill. Off by default, so behavior is unchanged unless the
        # user turns it on. (enabled stays non-persistent — see settings policy.)
        self.fill_blanks = bool(load_settings().get(FILL_BLANKS_KEY, False))
        self._stop_watcher = threading.Event()
        self._last_change_count = clipboard_change_count()
        self._base_icon = None  # cached full-color source image
        # Held while a help/welcome dialog is open so repeat clicks collapse
        # onto the one window instead of stacking modal boxes.
        self._dialog_lock = threading.Lock()
        # Keep watcher read/convert/write atomic relative to explicit exports.
        self._clipboard_operation_lock = threading.Lock()
        # Only one explicit formula read may be in flight.  Without this, a
        # slower worker for an older selection can overwrite a newer export.
        self._formula_export_lock = threading.Lock()
        # Launch-at-login via the MSIX StartupTask (parity with macOS). One read
        # seeds both flags; the state is None on source/dev runs and the bare
        # non-MSIX exe (no package identity), where the menu omits the toggle.
        login_state = startup_task.current_state()
        self.login_supported = login_state is not None
        self.login_enabled = login_state in startup_task.ENABLED_STATES
        self.icon = pystray.Icon(
            "Tabledown",
            icon=self._icon_image(),
            title="Tabledown",
            menu=self._build_menu(),
        )
        # Ctrl+Alt+T toggles conversion (parity with macOS ⌘⌃T). Created here,
        # started in run() — construction must stay side-effect-free for tests
        # that build the app without running it. MOD_NOREPEAT so holding the
        # keys flips once, not repeatedly.
        self._hotkey = GlobalHotkey(
            MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, VK_T, self._toggle_from_hotkey
        )
        # Ctrl+Alt+E runs the Excel table-with-formulas export action. It uses a
        # separate registration thread, so the fixed NULL-hwnd hotkey id remains
        # unique in each thread's message queue.
        self._formula_hotkey = GlobalHotkey(
            MOD_CONTROL | MOD_ALT | MOD_NOREPEAT,
            VK_E,
            self._copy_selected_excel_formulas_from_hotkey,
        )

    def run(self) -> None:
        watcher = threading.Thread(
            target=self._watch_clipboard,
            name="TabledownWindowsClipboardWatcher",
            daemon=True,
        )
        watcher.start()
        log("windows clipboard watcher started")
        if self._hotkey.start():
            log("windows global hotkey registered")
        else:
            log("windows global hotkey not registered")
        if self._formula_hotkey.start():
            log("windows Excel formula hotkey registered")
        else:
            log("windows Excel formula hotkey not registered")
        # setup runs on a pystray thread once the icon is ready; the first-run
        # welcome goes through _show_message_box_async, so it never blocks.
        self.icon.run(setup=self._on_ready)

    # --- Menu ---

    def _build_menu(self) -> pystray.Menu:
        language_items = [
            pystray.MenuItem(
                self._language_option_title(code),
                self._language_action(code),
            )
            for code in SUPPORTED_LANGUAGES
        ]
        items = [
            # Fixed label + checkmark (matching macOS): the old swapping label
            # ("Enabled ✓"/"Disabled") read as ambiguous — current state or
            # action? A checked, stable label is unambiguous.
            pystray.MenuItem(
                t("menu.toggle", self.lang),
                self.toggle,
                checked=lambda _item: self.enabled,
            ),
            pystray.MenuItem(
                t("menu.copy_excel_formulas", self.lang),
                self.copy_selected_excel_formulas,
            ),
            pystray.Menu.SEPARATOR,
            # Preference: fill merged-header blanks in the Markdown conversion
            # (0.5.0). Checkmark re-reads self.fill_blanks each time the menu opens.
            pystray.MenuItem(
                t("menu.fill_blanks", self.lang),
                self.toggle_fill_blanks,
                checked=lambda _item: self.fill_blanks,
            ),
            pystray.MenuItem(t("menu.language", self.lang), pystray.Menu(*language_items)),
        ]
        # Preferences (language + login) sit before help/quit. The login toggle
        # only appears where StartupTask is usable (packaged MSIX run).
        if self.login_supported:
            items.append(
                pystray.MenuItem(
                    t("menu.login_item", self.lang),
                    self.toggle_login_item,
                    checked=lambda _item: self.login_enabled,
                )
            )
        items.append(pystray.MenuItem(t("menu.diagnostics", self.lang), self.share_diagnostics))
        items.append(pystray.MenuItem(t("menu.help", self.lang), self.show_help))
        items.append(pystray.MenuItem(t("menu.quit", self.lang), self.quit_app))
        return pystray.Menu(*items)

    def _refresh_menu(self) -> None:
        self.icon.menu = self._build_menu()
        try:
            self.icon.update_menu()
        except NotImplementedError:
            pass

    def _language_action(self, code: str):
        # pystray._assert_action rejects callables whose positional arg count
        # isn't 0/1/2. The old `lambda _icon, _item, language=code` had THREE
        # positional params (the default-arg capture counts toward co_argcount),
        # so building this submenu raised ValueError and crashed the tray at
        # startup. A factory closure captures `code` while staying a 2-arg
        # callable that pystray accepts.
        def handler(_icon, _item):
            self._set_language(code)

        return handler

    def _language_option_title(self, code: str) -> str:
        mark = " ✓" if code == self.lang else ""
        return t(f"menu.language.{code}", self.lang) + mark

    def toggle(self, _icon, _item) -> None:
        # Menu click: runs on pystray's pump thread, so refreshing the menu is fine.
        self._set_enabled(not self.enabled, refresh_menu=True)

    def toggle_fill_blanks(self, _icon, _item) -> None:
        # Menu click (pump thread): flip, persist, and refresh so the checkmark
        # reflects the new state. Read live by the watcher on the next conversion.
        self.fill_blanks = not self.fill_blanks
        save_setting(FILL_BLANKS_KEY, self.fill_blanks)
        log(f"fill blanks {'enabled' if self.fill_blanks else 'disabled'}")
        self._refresh_menu()

    def _toggle_from_hotkey(self) -> None:
        # Ctrl+Alt+T: runs on the hotkey thread, NOT pystray's pump. The bool
        # flip is the source of truth the watcher reads, so the toggle takes
        # effect immediately regardless of any UI update. We skip _refresh_menu()
        # (rebuilding the Win32 HMENU off the pump thread is unsafe) — the
        # checkmark re-evaluates lazily from `self.enabled` the next time the
        # menu opens. The icon image IS repainted for immediate feedback (parity
        # with macOS, where the slash appears the instant you toggle); see the
        # tradeoff note in _set_enabled.
        self._set_enabled(not self.enabled, refresh_menu=False)

    def _set_enabled(self, value: bool, *, refresh_menu: bool) -> None:
        self.enabled = value
        if refresh_menu:
            self._refresh_menu()
        # Mirror the state on the icon itself (slash = off), like macOS.
        #
        # Tradeoff (hotkey path): `icon.icon =` is pystray's PUBLIC setter, but
        # its win32 impl does DestroyIcon→LoadImage→Shell_NotifyIcon on the
        # CALLING thread with no lock. From the hotkey thread that can race
        # pystray's own pump-thread icon writes (explorer restart →
        # WM_TASKBARCREATED, or a resolution change). The race is practically
        # unreachable — a single user cannot click the tray menu and press the
        # chord in the same microsecond, and explorer restarts are rare — and the
        # worst case is a transient wrong/blank tray icon that self-corrects on
        # the next repaint (a GDI handle misuse, never a crash). We accept it to
        # keep immediate feedback without the alternatives: a permanently stale
        # icon (skip the repaint) reads as "the toggle didn't work", and
        # marshalling onto the pump thread would couple to pystray internals
        # (_hwnd/_message_handlers) we can't exercise off Windows. The functional
        # toggle (the bool above) never depends on this repaint.
        try:
            self.icon.icon = self._icon_image()
        except Exception as exc:  # noqa: BLE001 - state flip must never crash the tray
            log(f"tray icon update failed: {exc}")
        log(f"conversion {'enabled' if self.enabled else 'disabled'}")

    def toggle_login_item(self, _icon, _item) -> None:
        # The WinRT toggle itself hops onto a private thread inside
        # startup_task; this callback (the pystray pump thread) only blocks
        # briefly on it, then refreshes the checkmark exactly like toggle().
        want = not self.login_enabled
        status = startup_task.set_enabled(want)
        self.login_enabled = status in startup_task.ENABLED_STATES
        log(f"login item toggle requested -> {want}, status={status}")
        self._refresh_menu()
        # Windows can refuse an enable the user/policy turned off elsewhere;
        # say so instead of leaving the checkmark silently unchecked.
        if want and status == "disabled_by_user":
            self._show_message_box_async(
                t("login_item.blocked_by_user", self.lang), t("help.title", self.lang)
            )
        elif want and status == "disabled_by_policy":
            self._show_message_box_async(
                t("login_item.blocked_by_policy", self.lang), t("help.title", self.lang)
            )

    def show_help(self, _icon, _item) -> None:
        self._show_message_box_async(
            t("help.message", self.lang), t("help.title", self.lang)
        )

    def copy_selected_excel_formulas(self, _icon, _item) -> None:
        """Export the active Excel table values and formulas without blocking."""

        if not self._formula_export_lock.acquire(blocking=False):
            return

        def worker() -> None:
            try:
                try:
                    result = read_selected_excel_formulas()
                except Exception as exc:  # noqa: BLE001 - guard daemon worker
                    log(f"excel formula read failed: {type(exc).__name__}")
                    self._show_message_box_async(
                        t("formula_export.error.export_failed", self.lang),
                        t("help.title", self.lang),
                    )
                    return
                if not result.ok:
                    key = f"formula_export.error.{result.code}"
                    message = t(key, self.lang)
                    if message == key:
                        safe_code = "unknown"
                        message = t("formula_export.error.export_failed", self.lang)
                    else:
                        safe_code = result.code
                    # Only a localized, known code reaches the shareable log.
                    # An unexpected adapter payload is reduced to "unknown".
                    log(f"excel formula export failed: {safe_code}")
                    self._show_message_box_async(message, t("help.title", self.lang))
                    return

                try:
                    xml = formula_selection_to_xml(result.selection)
                    with self._clipboard_operation_lock:
                        write_text_only_clipboard(xml)
                except Exception as exc:  # noqa: BLE001 - omit formula-bearing details
                    log(f"excel formula clipboard write failed: {type(exc).__name__}")
                    self._show_message_box_async(
                        t("formula_export.error.export_failed", self.lang),
                        t("help.title", self.lang),
                    )
                    return

                log("excel table-with-formulas export succeeded")
                self._show_message_box_async(
                    t("formula_export.success", self.lang), t("help.title", self.lang)
                )
            finally:
                self._formula_export_lock.release()

        try:
            threading.Thread(
                target=worker, name="TabledownWindowsExcelFormulaExport", daemon=True
            ).start()
        except Exception as exc:  # noqa: BLE001 - constructor/start failure must unlock
            self._formula_export_lock.release()
            log(f"excel formula worker failed to start: {type(exc).__name__}")
            self._show_message_box_async(
                t("formula_export.error.export_failed", self.lang),
                t("help.title", self.lang),
            )

    def _copy_selected_excel_formulas_from_hotkey(self) -> None:
        """Route Ctrl+Alt+E through the same action as the tray menu item."""
        self.copy_selected_excel_formulas(None, None)

    def share_diagnostics(self, _icon, _item) -> None:
        """Write a scrubbed local log and open its folder in Explorer.

        Local only — nothing is sent anywhere. Runs on its own thread so the
        export + Explorer launch never blocks pystray's message pump, and is
        fully wrapped: a pump-thread callback exception is caught by pystray,
        not by sys/threading.excepthook, so it must not escape.
        """
        def worker() -> None:
            try:
                path = diagnostics.export_diagnostics()
                if path is None:
                    # Do NOT reveal the raw, unscrubbed log as a "report".
                    self._show_message_box_async(
                        t("diagnostics.export_failed", self.lang), t("help.title", self.lang)
                    )
                    return
                diagnostics.reveal(path)
            except Exception as exc:  # noqa: BLE001 - never kill the tray
                log(f"share diagnostics failed: {type(exc).__name__}")

        try:
            threading.Thread(
                target=worker, name="TabledownWindowsDiagnostics", daemon=True
            ).start()
        except RuntimeError as exc:  # noqa: BLE001 - can't spawn (resource limit)
            log(f"diagnostics thread failed to start: {type(exc).__name__}")

    def quit_app(self, _icon, _item) -> None:
        self._stop_watcher.set()
        self._hotkey.stop()
        self._formula_hotkey.stop()
        self.icon.stop()

    def _set_language(self, code: str) -> None:
        if code == self.lang:
            return
        self.lang = code
        save_preferred_language(code)
        log(f"language set to {code}")
        self._refresh_menu()

    # --- First run ---

    def _on_ready(self, icon) -> None:
        icon.visible = True
        try:
            if load_settings().get(WELCOME_SHOWN_KEY):
                return
            # Marked before showing so a failing MessageBox can never loop the
            # welcome on every launch.
            save_setting(WELCOME_SHOWN_KEY, True)
            self._show_message_box_async(
                t("welcome.intro", self.lang) + "\n\n" + t("help.message", self.lang),
                t("welcome.title", self.lang),
            )
        except Exception as exc:  # noqa: BLE001 - welcome must never kill the tray
            log(f"welcome failed: {exc}")

    def _show_message_box_async(self, message: str, title: str) -> None:
        """Show a modal info box without blocking the caller.

        Menu callbacks run synchronously on pystray's message-pump thread, so a
        blocking MessageBox there freezes the tray and — with no foreground
        rights — can open behind the active window. The user then clicks Help
        again, stacking boxes that look like they "won't close". Running it on
        its own thread keeps the pump free; the lock collapses repeat clicks
        onto the single open dialog.
        """
        if not self._dialog_lock.acquire(blocking=False):
            return

        def worker() -> None:
            try:
                self._message_box(message, title)
            finally:
                self._dialog_lock.release()

        try:
            threading.Thread(
                target=worker, name="TabledownWindowsDialog", daemon=True
            ).start()
        except RuntimeError as exc:  # noqa: BLE001 - can't spawn (resource limit)
            # Release the lock we just took so a later click can still try.
            self._dialog_lock.release()
            log(f"dialog thread failed to start: {exc}")

    @staticmethod
    def _message_box(message: str, title: str) -> None:
        flags = _MB_ICONINFORMATION | _MB_SETFOREGROUND | _MB_TOPMOST
        ctypes.windll.user32.MessageBoxW(None, message, title, flags)

    # --- Clipboard watcher ---

    def _watch_clipboard(self) -> None:
        while not self._stop_watcher.is_set():
            time.sleep(0.1)
            if not self.enabled:
                continue

            current_change_count = clipboard_change_count()
            if current_change_count == self._last_change_count:
                continue

            self._last_change_count = current_change_count
            self._augment_clipboard()

    def _augment_clipboard(self) -> None:
        try:
            with self._clipboard_operation_lock:
                content = read_clipboard()
                updated = converted_clipboard(content, self.fill_blanks)
                if updated is None:
                    return

                write_clipboard(
                    text=updated.get("text"),
                    html=updated.get("html"),
                    mark_generated=True,
                    drop_formats=updated.get("drop_formats"),
                )
            log("clipboard formats updated")
        except Exception as exc:  # noqa: BLE001 - tray app should keep watching
            # Type only — this processes clipboard content and exc messages can
            # carry table data we must never write to the shareable log.
            log(f"clipboard update failed: {type(exc).__name__}")

    # --- Icon ---

    def _icon_image(self) -> Image.Image:
        """Tray icon for the current state — a slash overlay means off."""
        image = self._base_image().copy()
        if not self.enabled:
            self._draw_off_slash(image)
        return image

    def _base_image(self) -> Image.Image:
        if self._base_icon is None:
            size = self._tray_icon_size()
            image = self._load_asset_icon(size)
            if image is None:
                image = self._draw_fallback_icon(size)
            self._base_icon = image
        return self._base_icon

    @staticmethod
    def _tray_icon_size() -> int:
        """The exact small-icon size the tray renders at (DPI-scaled).

        Matching it avoids any downstream scaling — the old fixed 64px source
        was rescaled by Windows to 16-24px and came out soft.
        """
        try:
            size = int(ctypes.windll.user32.GetSystemMetrics(_SM_CXSMICON))
            if size >= 16:
                return size
        except Exception:  # noqa: BLE001 - e.g. non-Windows test environment
            pass
        return 64

    def _load_asset_icon(self, size: int) -> Image.Image | None:
        candidates = []
        if getattr(sys, "frozen", False):
            candidates.append(Path(getattr(sys, "_MEIPASS", "")) / "assets" / "generated" / "tablemark_app_1024.png")
        module_path = Path(__file__).resolve()
        candidates.append(module_path.parents[2] / "assets" / "generated" / "tablemark_app_1024.png")

        for path in candidates:
            if not path.exists():
                continue
            try:
                return Image.open(path).convert("RGBA").resize((size, size), Image.LANCZOS)
            except OSError:
                continue
        return None

    @staticmethod
    def _draw_fallback_icon(size: int) -> Image.Image:
        image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        s = size / 64  # original geometry was authored on a 64px canvas
        line = max(1, round(3 * s))
        draw.rounded_rectangle(
            (round(8 * s), round(8 * s), round(56 * s), round(56 * s)),
            radius=max(2, round(10 * s)),
            fill=(18, 24, 38, 255),
        )
        for x in (24, 40):
            draw.line((round(x * s), round(12 * s), round(x * s), round(52 * s)),
                      fill=(255, 255, 255, 230), width=line)
        for y in (24, 40):
            draw.line((round(12 * s), round(y * s), round(52 * s), round(y * s)),
                      fill=(255, 255, 255, 230), width=line)
        return image

    @staticmethod
    def _draw_off_slash(image: Image.Image) -> None:
        """Red diagonal slash with a punched gap — the macOS off convention."""
        size = image.size[0]
        margin = max(1, round(size * 0.08))
        punch_width = max(3, round(size * 0.22))
        slash_width = max(2, round(size * 0.11))
        ends = (margin, size - margin, size - margin, margin)
        punch = Image.new("L", image.size, 0)
        ImageDraw.Draw(punch).line(ends, fill=255, width=punch_width)
        image.putalpha(ImageChops.subtract(image.getchannel("A"), punch))
        ImageDraw.Draw(image).line(ends, fill=(214, 60, 60, 255), width=slash_width)


def main() -> None:
    # Refuse a second tray instance (manual launch on top of the StartupTask, a
    # stale copy a crash left behind, …). Two clipboard watchers racing would
    # stomp each other's conversions. macOS gets this from LaunchServices.
    if not single_instance.acquire_single_instance():
        return
    # Capture failures that otherwise vanish (watcher-thread exceptions, native
    # faults) into the local log — purely local, nothing is transmitted.
    diagnostics.install_crash_hooks()
    TabledownWindowsApp().run()


if __name__ == "__main__":
    main()
