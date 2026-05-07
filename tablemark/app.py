"""Tabledown - Excel ↔ Markdown table converter for macOS."""
from pathlib import Path
import threading
import time

import rumps

from .clipboard import clipboard_change_count, read_clipboard, write_clipboard
from .converter.html_to_md import html_table_to_markdown
from .converter.md_to_tsv import (
    is_markdown_table,
    markdown_table_to_html,
)
from .logger import log


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
        self.toggle_item = rumps.MenuItem("활성화 ✓", callback=self.toggle)
        self.menu = [
            self.toggle_item,
            None,  # separator
            rumps.MenuItem("도움말", callback=self.show_help),
            rumps.MenuItem("종료", callback=self.quit_app),
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

    # --- Menu actions ---

    def toggle(self, sender):
        self.enabled = not self.enabled
        sender.title = "활성화 ✓" if self.enabled else "비활성화"

    def show_help(self, _):
        rumps.alert(
            title="Tabledown",
            message=(
                "Excel ↔ Markdown 표 변환기\n\n"
                "사용법:\n"
                "1. Excel/스프레드시트 또는 마크다운 표를 복사 (Cmd+C)\n"
                "2. 원하는 앱에서 그대로 붙여넣기 (Cmd+V)\n\n"
                "Excel 표를 복사하면 마크다운 에디터에서 Markdown 표로 붙고,\n"
                "Markdown 표를 복사하면 Excel에서 셀에 분리되어 붙습니다."
            ),
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

            write_clipboard(**updated)
            self._last_change_count = clipboard_change_count()
            log("clipboard formats updated")
        except Exception as e:
            log(f"clipboard update failed: {e}")

    def _converted_clipboard(self, content):
        """Return clipboard formats to write, or None when no update is needed."""
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
                "text": markdown,
                "html": html,
            }

        return None

def main():
    TabledownApp().run()


if __name__ == "__main__":
    main()
