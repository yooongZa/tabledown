"""Tabledown - Excel ↔ Markdown table converter for macOS."""
from pathlib import Path
import threading
import time

import rumps

from .clipboard import (
    HTML_TYPES,
    RENDERED_TABLE_TYPES,
    clipboard_change_count,
    read_clipboard,
    write_clipboard,
)
from .converter.html_to_md import html_table_to_markdown
from .converter.md_to_tsv import (
    is_markdown_table,
    markdown_table_to_html,
)
from .i18n import (
    SUPPORTED_LANGUAGES,
    resolve_language,
    save_preferred_language,
    t,
)
from .logger import log


def _markdown_paste_block(markdown: str) -> str:
    """Return markdown padded so block parsers see a standalone table."""
    return "\n" + markdown.strip() + "\n"


class TabledownApp(rumps.App):
    ICON_NAME = "tablemark_menu_40.png"

    def __init__(self):
        log("app starting")
        super().__init__(
            "Tabledown",
            icon=self._icon_path(),
            template=True,
            quit_button=None,
        )
        self.enabled = True
        self.lang = resolve_language()
        log(f"language resolved to {self.lang}")

        self.toggle_item = rumps.MenuItem(
            self._toggle_title(),
            callback=self.toggle,
        )
        self.language_item = rumps.MenuItem(t("menu.language", self.lang))
        self.language_options = {}
        for code in SUPPORTED_LANGUAGES:
            item = rumps.MenuItem(
                self._language_option_title(code),
                callback=self._make_language_setter(code),
            )
            self.language_options[code] = item
        self.language_item.update(list(self.language_options.values()))

        self.help_item = rumps.MenuItem(t("menu.help", self.lang), callback=self.show_help)
        self.quit_item = rumps.MenuItem(t("menu.quit", self.lang), callback=self.quit_app)

        self.menu = [
            self.toggle_item,
            None,  # separator
            self.language_item,
            self.help_item,
            self.quit_item,
        ]

        self._stop_watcher = threading.Event()
        self._last_change_count = clipboard_change_count()
        self._watcher_thread = threading.Thread(
            target=self._watch_clipboard,
            name="TabledownClipboardWatcher",
            daemon=True,
        )
        self._watcher_thread.start()
        log("clipboard watcher started")

    @classmethod
    def _icon_path(cls):
        module_path = Path(__file__).resolve()
        candidates = [
            module_path.parent.parent / "assets" / "generated" / cls.ICON_NAME,
            module_path.parent / cls.ICON_NAME,
        ]
        candidates.extend(parent / "Resources" / cls.ICON_NAME for parent in module_path.parents)
        for path in candidates:
            if path.exists():
                return str(path)
        log("menu icon not found; falling back to title")
        return None

    # --- Localization ---

    def _toggle_title(self) -> str:
        return t("menu.toggle_on" if self.enabled else "menu.toggle_off", self.lang)

    def _language_option_title(self, code: str) -> str:
        mark = " ✓" if code == self.lang else ""
        return t(f"menu.language.{code}", self.lang) + mark

    def _apply_language(self):
        """Refresh every menu title for the current language."""
        self.toggle_item.title = self._toggle_title()
        self.language_item.title = t("menu.language", self.lang)
        for code, item in self.language_options.items():
            item.title = self._language_option_title(code)
        self.help_item.title = t("menu.help", self.lang)
        self.quit_item.title = t("menu.quit", self.lang)

    def _make_language_setter(self, code: str):
        def _set_language(_sender):
            if self.lang == code:
                return
            self.lang = code
            save_preferred_language(code)
            log(f"language set to {code}")
            self._apply_language()
        return _set_language

    # --- Menu actions ---

    def toggle(self, _sender):
        self.enabled = not self.enabled
        self.toggle_item.title = self._toggle_title()

    def show_help(self, _):
        rumps.alert(
            title=t("help.title", self.lang),
            message=t("help.message", self.lang),
        )

    def quit_app(self, _):
        self._stop_watcher.set()
        rumps.quit_application()

    # --- Clipboard watcher ---

    def _watch_clipboard(self):
        while not self._stop_watcher.is_set():
            time.sleep(0.1)
            if not self.enabled:
                continue

            current_change_count = clipboard_change_count()
            if current_change_count == self._last_change_count:
                continue

            self._last_change_count = current_change_count
            self._augment_clipboard()

    def _augment_clipboard(self):
        try:
            content = read_clipboard()
            updated = self._converted_clipboard(content)
            if updated is None:
                return

            write_clipboard(**updated, mark_generated=True)
            log("clipboard formats updated")
        except Exception as e:
            log(f"clipboard update failed: {e}")

    def _converted_clipboard(self, content):
        """Return clipboard formats to write, or None when no update is needed."""
        if content.get("generated"):
            return None

        html = content.get("html", "")
        text = content.get("text", "")

        if text and is_markdown_table(text):
            if html and "<table" in html.lower():
                return None
            log("detected markdown table")
            return {
                "text": text,
                "html": markdown_table_to_html(text),
            }

        if html and "<table" in html.lower():
            markdown = html_table_to_markdown(html)
            if text.strip() == markdown.strip():
                return None
            log("detected html table")
            return {
                "text": _markdown_paste_block(markdown),
                "drop_types": RENDERED_TABLE_TYPES | HTML_TYPES,
            }

        return None

def main():
    TabledownApp().run()


if __name__ == "__main__":
    main()
