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
from .converter.html_to_md import (
    convert_document_tables,
    forward_fill_key_columns,
    html_has_content_outside_table,
    html_table_to_markdown,
    html_table_to_model,
)
from .converter.md_to_tsv import (
    is_markdown_table,
    markdown_table_to_html,
    markdown_table_to_rows,
)
from .converter.table_xml import (
    is_table_xml,
    model_to_xml,
    table_xml_to_model,
)
from .i18n import (
    SUPPORTED_LANGUAGES,
    resolve_language,
    save_preferred_language,
    t,
)
from . import login_item
from .hotkey import GlobalHotkey
from .logger import log
from .settings import load_fill_blanks, save_fill_blanks
from .store import Store


def _markdown_paste_block(markdown: str) -> str:
    """Return markdown padded so block parsers see a standalone table."""
    return "\n" + markdown.strip() + "\n"


class TabledownApp(rumps.App):
    ICON_NAME = "tablemark_menu_40.png"
    ICON_NAME_OFF = "tablemark_menu_40_off.png"

    def __init__(self):
        log("app starting")
        self._clear_stale_status_item_visibility()
        self.enabled = True
        self.fill_blanks = load_fill_blanks()
        super().__init__(
            "Tabledown",
            icon=self._current_icon_path(),
            template=True,
            quit_button=None,
        )
        self.lang = resolve_language()
        log(f"language resolved to {self.lang}")

        self.toggle_item = rumps.MenuItem(
            t("menu.toggle", self.lang),
            callback=self.toggle,
        )
        self.toggle_item.state = 1 if self.enabled else 0
        self.copy_xml_item = rumps.MenuItem(
            t("menu.copy_xml", self.lang),
            callback=self.copy_as_xml,
            key="c",
        )
        # Show "⌘⌃C" next to the item. rumps' `key` gives a Command-only
        # equivalent; OR in Control so the displayed/registered accelerator
        # matches the global Carbon hotkey (⌘⌃C). Best-effort — purely cosmetic
        # discoverability, so never let it crash startup.
        try:
            from AppKit import (
                NSCommandKeyMask,
                NSControlKeyMask,
            )
            self.copy_xml_item._menuitem.setKeyEquivalentModifierMask_(
                NSCommandKeyMask | NSControlKeyMask
            )
        except Exception as exc:  # noqa: BLE001
            log(f"failed to set XML shortcut modifier mask: {exc}")
        self.fill_blanks_item = rumps.MenuItem(
            t("menu.fill_blanks", self.lang),
            callback=self.toggle_fill_blanks,
        )
        self.fill_blanks_item.state = 1 if self.fill_blanks else 0
        self.fill_blanks_item._menuitem.setToolTip_(t("menu.fill_blanks_tooltip", self.lang))
        self.language_item = rumps.MenuItem(t("menu.language", self.lang))
        self.language_options = {}
        for code in SUPPORTED_LANGUAGES:
            item = rumps.MenuItem(
                self._language_option_title(code),
                callback=self._make_language_setter(code),
            )
            item.state = 1 if code == self.lang else 0
            self.language_options[code] = item
        self.language_item.update(list(self.language_options.values()))

        # --- In-App Purchases (donations + XML Pro subscription) ---
        # Strong-ref the store on self: it owns the StoreKit transaction observer
        # whose lifetime must match the app's. Fetch product metadata now so the
        # menu prices/titles can populate before the user opens a purchase sheet.
        self.store = Store()
        self.store.on_tip_purchased = self._on_tip_purchased
        self.store.fetch_products()

        self.donate_item = rumps.MenuItem(t("menu.donate", self.lang))
        self.donate_options = {}
        for tier in ("small", "medium", "large"):
            item = rumps.MenuItem(
                t(f"menu.donate.{tier}", self.lang),
                callback=self._make_tip_handler(tier),
            )
            self.donate_options[tier] = item
        self.donate_item.update(list(self.donate_options.values()))

        self.restore_item = rumps.MenuItem(
            t("menu.restore", self.lang),
            callback=self.restore_purchases,
        )

        self.login_item_supported = login_item.is_supported()
        if self.login_item_supported:
            self.login_item_menu = rumps.MenuItem(
                t("menu.login_item", self.lang),
                callback=self.toggle_login_item,
            )
            self.login_item_menu.state = 1 if login_item.is_enabled() else 0
        else:
            self.login_item_menu = None

        self.help_item = rumps.MenuItem(t("menu.help", self.lang), callback=self.show_help)
        self.quit_item = rumps.MenuItem(t("menu.quit", self.lang), callback=self.quit_app)

        self.settings_item = rumps.MenuItem(t("menu.settings", self.lang))
        settings_children = [self.fill_blanks_item, self.language_item]
        if self.login_item_menu is not None:
            settings_children.append(self.login_item_menu)
        settings_children.append(self.restore_item)  # Apple-required Restore
        self.settings_item.update(settings_children)

        self.menu = [
            self.toggle_item,
            self.copy_xml_item,
            None,  # separator
            self.donate_item,
            self.settings_item,
            None,  # separator
            self.help_item,
            self.quit_item,
        ]

        # --- Global hotkey ⌘⌃C -> copy_as_xml (gated inside copy_as_xml) ---
        # Strong-ref on self so the Carbon callback trampoline survives GC.
        # Graceful: if registration fails the menu item still works.
        self.hotkey = GlobalHotkey(self.copy_as_xml)
        self.hotkey.register()

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
    def _icon_path(cls, name: str):
        module_path = Path(__file__).resolve()
        candidates = [
            module_path.parent.parent / "assets" / "generated" / name,
            module_path.parent / name,
        ]
        candidates.extend(parent / "Resources" / name for parent in module_path.parents)
        for path in candidates:
            if path.exists():
                return str(path)
        return None

    def _current_icon_path(self):
        name = self.ICON_NAME if self.enabled else self.ICON_NAME_OFF
        path = self._icon_path(name)
        if path is None and not self.enabled:
            # Off variant missing — fall back to the regular icon so the menu
            # never becomes a bare text title.
            path = self._icon_path(self.ICON_NAME)
        if path is None:
            log("menu icon not found; falling back to title")
        return path

    @staticmethod
    def _clear_stale_status_item_visibility():
        """Strip leftover NSStatusItem visibility flags from NSUserDefaults.

        Earlier 0.1.0 builds shipped a "hide menu bar icon" action that flipped
        NSStatusItem's autosave-backed visibility flag to 0. Once set, the
        LSUIElement app started invisible on every launch, leaving no UI to
        bring it back. The action is gone in 0.1.1, but lingering "NSStatusItem
        Visible*" keys would still hide the icon on first run of the new build.
        Wiping any such key on startup makes the icon reappear, and because the
        hide action no longer exists, autosave will only ever record the
        visible state from here on.
        """
        try:
            from Foundation import NSUserDefaults
            defaults = NSUserDefaults.standardUserDefaults()
            stale = [
                str(key)
                for key in defaults.dictionaryRepresentation().keys()
                if str(key).startswith("NSStatusItem Visible")
            ]
            for key in stale:
                defaults.removeObjectForKey_(key)
                log(f"removed stale visibility key: {key}")
            if stale:
                defaults.synchronize()
        except Exception as exc:
            log(f"failed to clear stale visibility defaults: {exc}")

    # --- Localization ---

    def _language_option_title(self, code: str) -> str:
        return t(f"menu.language.{code}", self.lang)

    def _apply_language(self):
        """Refresh every menu title for the current language."""
        self.toggle_item.title = t("menu.toggle", self.lang)
        self.copy_xml_item.title = t("menu.copy_xml", self.lang)
        self.donate_item.title = t("menu.donate", self.lang)
        for tier, item in self.donate_options.items():
            item.title = t(f"menu.donate.{tier}", self.lang)
        self.restore_item.title = t("menu.restore", self.lang)
        self.fill_blanks_item.title = t("menu.fill_blanks", self.lang)
        self.fill_blanks_item._menuitem.setToolTip_(t("menu.fill_blanks_tooltip", self.lang))
        self.settings_item.title = t("menu.settings", self.lang)
        self.language_item.title = t("menu.language", self.lang)
        for code, item in self.language_options.items():
            item.title = self._language_option_title(code)
        if self.login_item_menu is not None:
            self.login_item_menu.title = t("menu.login_item", self.lang)
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
            for c, item in self.language_options.items():
                item.state = 1 if c == code else 0
        return _set_language

    # --- Menu actions ---

    def toggle(self, _sender):
        self.enabled = not self.enabled
        self.toggle_item.state = 1 if self.enabled else 0
        self.icon = self._current_icon_path()
        log(f"conversion {'enabled' if self.enabled else 'disabled'}")

    def toggle_login_item(self, _sender):
        new_state = not login_item.is_enabled()
        login_item.set_enabled(new_state)
        actual = login_item.is_enabled()
        log(f"login item toggle requested -> {new_state}, actual={actual}")
        if self.login_item_menu is not None:
            self.login_item_menu.state = 1 if actual else 0

    def toggle_fill_blanks(self, _sender):
        self.fill_blanks = not self.fill_blanks
        self.fill_blanks_item.state = 1 if self.fill_blanks else 0
        save_fill_blanks(self.fill_blanks)
        log(f"fill blanks {'enabled' if self.fill_blanks else 'disabled'}")

    def copy_as_xml(self, _sender):
        """Convert the table currently on the clipboard to LLM-friendly XML.

        This is a deliberate user action: it puts XML in the text slot and drops
        the HTML <table> so a paste anywhere yields the XML the user asked for.
        There is no automatic XML→table direction — XML is produced only by this
        menu click, never inferred from clipboard contents by the watcher.
        """
        # XML is a Pro feature: gate the user-action path (menu click + ⌘⌃C
        # hotkey both land here) before any conversion. This is the ONLY gating
        # point — the watcher (_converted_clipboard) and converter functions are
        # untouched, per the clipboard invariants in CLAUDE.md. Everything below
        # this guard is unchanged.
        if not self.store.pro_active:
            self._present_subscription_sheet()
            return
        try:
            model = self._clipboard_table_model(read_clipboard(), self.fill_blanks)
            if model is None:
                rumps.alert(
                    title=t("xml.no_table_title", self.lang),
                    message=t("xml.no_table_message", self.lang),
                )
                return
            header_levels, data_rows = model
            write_clipboard(
                text=model_to_xml(header_levels, data_rows),
                mark_generated=True,
                drop_types=RENDERED_TABLE_TYPES | HTML_TYPES,
            )
            log("copied clipboard table as XML")
        except Exception as exc:  # noqa: BLE001 - surface failure to the user
            log(f"copy as xml failed: {exc}")
            rumps.alert(
                title=t("xml.no_table_title", self.lang),
                message=t("xml.no_table_message", self.lang),
            )

    # --- In-App Purchase actions ---

    def _present_subscription_sheet(self):
        """Show the Pro subscription prompt and start checkout if accepted.

        Uses rumps.alert with an OK (subscribe) / Cancel choice. rumps.alert
        returns 1 for the default (OK) button. Triggered only from the gated
        copy_as_xml path when the user lacks the Pro entitlement.
        """
        try:
            response = rumps.alert(
                title=t("xml.locked_title", self.lang),
                message=t("xml.locked_message", self.lang),
                ok=t("xml.subscribe_button", self.lang),
                cancel=True,
            )
        except Exception as exc:  # noqa: BLE001
            log(f"subscription sheet failed: {exc}")
            return
        if response == 1:
            log("user chose to subscribe to Pro")
            self.store.subscribe_pro()

    def _make_tip_handler(self, tier: str):
        def _handler(_sender):
            log(f"donation requested: {tier}")
            self.store.purchase_tip(tier)
        return _handler

    def restore_purchases(self, _sender):
        log("restore purchases requested")
        self.store.restore()

    def _on_tip_purchased(self, _product_id):
        """Store callback: thank the user after a donation completes."""
        try:
            rumps.alert(
                title=t("help.title", self.lang),
                message=t("tip.thanks", self.lang),
            )
        except Exception as exc:  # noqa: BLE001
            log(f"thanks alert failed: {exc}")

    @staticmethod
    def _clipboard_table_model(content, fill_blanks=False):
        """Extract a ``(header_levels, data_rows)`` table model from clipboard, or None.

        Accepts an HTML <table> (Excel/Sheets), table XML, or a Markdown table,
        in that priority order. ``header_levels`` is one list per header level so
        multi-level group headers survive (a Markdown/plain table has a single
        level). When ``fill_blanks`` is set, blank cells in the left grouping
        columns are forward-filled (see forward_fill_key_columns) — the
        user-controlled "XML: 빈칸을 자동 채우기" option, off by default.
        """
        html = content.get("html", "")
        text = content.get("text", "")
        model = None
        if html and "<table" in html.lower():
            try:
                # Merge-aware: forward-fills rowspan, skips a full-width title
                # row, and keeps multi-level headers — see html_table_to_model.
                model = html_table_to_model(html)
            except ValueError:
                model = None
        if model is None and text and is_table_xml(text):
            try:
                model = table_xml_to_model(text)
            except ValueError:
                model = None
        if model is None and text and is_markdown_table(text, strict=False):
            try:
                rows = markdown_table_to_rows(text)
                model = ([rows[0]], rows[1:])
            except (ValueError, IndexError):
                model = None
        if model is None:
            return None
        header_levels, data_rows = model
        if not data_rows:
            return None
        if fill_blanks:
            # forward_fill operates on flat rows (header at [0], never filled);
            # pass the deepest header level as the placeholder, take data back.
            data_rows = forward_fill_key_columns([header_levels[-1]] + data_rows)[1:]
        return header_levels, data_rows

    def show_help(self, _):
        rumps.alert(
            title=t("help.title", self.lang),
            message=t("help.message", self.lang),
        )

    def quit_app(self, _):
        self._stop_watcher.set()
        if getattr(self, "hotkey", None) is not None:
            self.hotkey.unregister()
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
        """Return clipboard formats to write, or None when no update is needed.

        The branch logic here encodes hard-won clipboard invariants — see the
        "클립보드 변환 불변식 (회귀 금지)" section in CLAUDE.md before changing
        it. In short: never strip the HTML slot for web tables or documents,
        and only apply the strict cell-count check to html-less text.
        """
        if content.get("generated"):
            return None

        html = content.get("html", "")
        text = content.get("text", "")
        has_html_table = bool(html) and "<table" in html.lower()

        # An accompanying HTML <table> is independent proof of a real table, so
        # skip the strict cell-count check (which only guards html-less text
        # that merely looks like a table). text+html is already multi-format —
        # leave it untouched so the destination picks its own format
        # (Excel -> table, markdown editor -> text).
        if text and is_markdown_table(text, strict=not has_html_table):
            if has_html_table:
                return None
            log("detected markdown table")
            return {
                "text": text,
                "html": markdown_table_to_html(text),
            }

        if has_html_table:
            # A table embedded in a document (headings, paragraphs, lists):
            # render its tables as Markdown in the text slot but KEEP the
            # original HTML slot. A Markdown editor reading text gets Markdown
            # tables; rich editors (Word, Excel) still read the original
            # <table> from HTML and paste a real table. Only RENDERED image
            # formats are dropped, never the HTML.
            if html_has_content_outside_table(html):
                converted = convert_document_tables(html)
                if not converted.strip() or converted.strip() == text.strip():
                    return None
                log("detected table in document")
                return {
                    "text": converted,
                    "drop_types": RENDERED_TABLE_TYPES,
                }
            # A bare table (Excel/Sheets): put a Markdown table in the text slot
            # but KEEP the original HTML <table>. A Markdown editor reading text
            # gets Markdown; Excel/Word reading HTML still paste a real table, so
            # one copy works for every destination (same rule as web tables and
            # documents). Only RENDERED image formats are dropped, never HTML.
            markdown = html_table_to_markdown(html)
            if text.strip() == markdown.strip():
                return None
            log("detected html table")
            return {
                "text": _markdown_paste_block(markdown),
                "drop_types": RENDERED_TABLE_TYPES,
            }

        return None

def main():
    TabledownApp().run()


if __name__ == "__main__":
    main()
