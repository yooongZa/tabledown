"""Windows tray app for Tabledown."""
from __future__ import annotations

import ctypes
from pathlib import Path
import sys
import threading
import time

from PIL import Image, ImageChops, ImageDraw
import pystray

from .conversion import converted_clipboard
from .i18n import SUPPORTED_LANGUAGES, resolve_language, save_preferred_language, t
from .logger import log
from .settings import load_settings, save_setting
from .win_clipboard import clipboard_change_count, read_clipboard, write_clipboard

WELCOME_SHOWN_KEY = "welcome_shown"

# MessageBoxW flags: MB_ICONINFORMATION.
_MB_ICONINFORMATION = 0x00000040
# GetSystemMetrics index for the small-icon width the tray actually uses
# (DPI-scaled by Windows: 16 at 96dpi, larger on HiDPI).
_SM_CXSMICON = 49


class TabledownWindowsApp:
    """Windows notification area app."""

    def __init__(self):
        log("windows app starting")
        self.enabled = True
        self.lang = resolve_language()
        self._stop_watcher = threading.Event()
        self._last_change_count = clipboard_change_count()
        self._base_icon = None  # cached full-color source image
        self.icon = pystray.Icon(
            "Tabledown",
            icon=self._icon_image(),
            title="Tabledown",
            menu=self._build_menu(),
        )

    def run(self) -> None:
        watcher = threading.Thread(
            target=self._watch_clipboard,
            name="TabledownWindowsClipboardWatcher",
            daemon=True,
        )
        watcher.start()
        log("windows clipboard watcher started")
        # setup runs on a pystray thread once the icon is ready; the first-run
        # MessageBox may block that thread without holding up the tray itself.
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
        return pystray.Menu(
            # Fixed label + checkmark (matching macOS): the old swapping label
            # ("Enabled ✓"/"Disabled") read as ambiguous — current state or
            # action? A checked, stable label is unambiguous.
            pystray.MenuItem(
                t("menu.toggle", self.lang),
                self.toggle,
                checked=lambda _item: self.enabled,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(t("menu.language", self.lang), pystray.Menu(*language_items)),
            pystray.MenuItem(t("menu.help", self.lang), self.show_help),
            pystray.MenuItem(t("menu.quit", self.lang), self.quit_app),
        )

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
        self.enabled = not self.enabled
        self._refresh_menu()
        # Mirror the state on the icon itself (slash = off), like macOS.
        try:
            self.icon.icon = self._icon_image()
        except Exception as exc:  # noqa: BLE001 - state flip must never crash the tray
            log(f"tray icon update failed: {exc}")

    def show_help(self, _icon, _item) -> None:
        self._message_box(t("help.message", self.lang), t("help.title", self.lang))

    def quit_app(self, _icon, _item) -> None:
        self._stop_watcher.set()
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
            self._message_box(
                t("welcome.intro", self.lang) + "\n\n" + t("help.message", self.lang),
                t("welcome.title", self.lang),
            )
        except Exception as exc:  # noqa: BLE001 - welcome must never kill the tray
            log(f"welcome failed: {exc}")

    @staticmethod
    def _message_box(message: str, title: str) -> None:
        ctypes.windll.user32.MessageBoxW(None, message, title, _MB_ICONINFORMATION)

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
            content = read_clipboard()
            updated = converted_clipboard(content)
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
            log(f"clipboard update failed: {exc}")

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
    TabledownWindowsApp().run()


if __name__ == "__main__":
    main()
