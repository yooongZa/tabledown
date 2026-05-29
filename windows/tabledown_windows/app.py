"""Windows tray app for Tabledown."""
from __future__ import annotations

import ctypes
from pathlib import Path
import sys
import threading
import time

from PIL import Image, ImageDraw
import pystray

from .conversion import converted_clipboard
from .i18n import SUPPORTED_LANGUAGES, resolve_language, save_preferred_language, t
from .logger import log
from .win_clipboard import clipboard_change_count, read_clipboard, write_clipboard


class TabledownWindowsApp:
    """Windows notification area app."""

    def __init__(self):
        log("windows app starting")
        self.enabled = True
        self.lang = resolve_language()
        self._stop_watcher = threading.Event()
        self._last_change_count = clipboard_change_count()
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
        self.icon.run()

    # --- Menu ---

    def _build_menu(self) -> pystray.Menu:
        language_items = [
            pystray.MenuItem(
                self._language_option_title(code),
                lambda _icon, _item, language=code: self._set_language(language),
            )
            for code in SUPPORTED_LANGUAGES
        ]
        return pystray.Menu(
            pystray.MenuItem(self._toggle_title(), self.toggle),
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

    def _toggle_title(self) -> str:
        return t("menu.toggle_on" if self.enabled else "menu.toggle_off", self.lang)

    def _language_option_title(self, code: str) -> str:
        mark = " ✓" if code == self.lang else ""
        return t(f"menu.language.{code}", self.lang) + mark

    def toggle(self, _icon, _item) -> None:
        self.enabled = not self.enabled
        self._refresh_menu()

    def show_help(self, _icon, _item) -> None:
        ctypes.windll.user32.MessageBoxW(
            None,
            t("help.message", self.lang),
            t("help.title", self.lang),
            0x00000040,
        )

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
        image = self._load_asset_icon()
        if image is not None:
            return image
        return self._draw_fallback_icon()

    def _load_asset_icon(self) -> Image.Image | None:
        candidates = []
        if getattr(sys, "frozen", False):
            candidates.append(Path(getattr(sys, "_MEIPASS", "")) / "assets" / "generated" / "tablemark_app_1024.png")
        module_path = Path(__file__).resolve()
        candidates.append(module_path.parents[2] / "assets" / "generated" / "tablemark_app_1024.png")

        for path in candidates:
            if not path.exists():
                continue
            try:
                return Image.open(path).convert("RGBA").resize((64, 64), Image.LANCZOS)
            except OSError:
                continue
        return None

    def _draw_fallback_icon(self) -> Image.Image:
        image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((8, 8, 56, 56), radius=10, fill=(18, 24, 38, 255))
        for x in (24, 40):
            draw.line((x, 12, x, 52), fill=(255, 255, 255, 230), width=3)
        for y in (24, 40):
            draw.line((12, y, 52, y), fill=(255, 255, 255, 230), width=3)
        return image


def main() -> None:
    TabledownWindowsApp().run()


if __name__ == "__main__":
    main()
