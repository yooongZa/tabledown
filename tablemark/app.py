"""Tabledown - Excel ↔ Markdown table converter for macOS."""
from pathlib import Path
import threading
import time

import rumps
from PyObjCTools import AppHelper

from .clipboard import (
    ClipboardChangedError,
    ClipboardWriteError,
    HTML_TYPES,
    RENDERED_TABLE_TYPES,
    clipboard_change_count,
    read_clipboard,
    write_clipboard,
    write_text_only_clipboard,
    write_table_clipboard,
)
from .converter.formula_export import (
    FormulaXmlTooLargeError,
    formula_copy_notice_key,
    formula_selection_to_ai_xml,
    MAX_XML_BYTES,
)
from .converter.html_to_md import (
    MultipleTablesError,
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
    TableXmlTooLargeError,
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
from .excel_formula import ExcelFormulaError, read_stable_selected_excel_formulas
from .excel_table import (
    excel_table_selection_to_html,
    excel_table_selection_to_model,
    excel_table_selection_to_model_with_sources,
    excel_table_selection_xml_metadata,
    read_stable_selected_excel_table,
)
from . import diagnostics
from . import login_item
from .hotkey import CMD, CONTROL, KEY_E, KEY_T, KEY_X, GlobalHotkeys
from .logger import log
from . import __version__
from .settings import (
    load_fill_blanks,
    load_welcome_shown,
    save_fill_blanks,
    save_welcome_shown,
)


GITHUB_URL = "https://github.com/yooongZa/tabledown"
_EXPORT_TABLE = "table"
_EXPORT_FORMULAS = "formulas"
_EXPORT_MARKDOWN = "markdown"
MAX_MARKDOWN_EXPORT_BYTES = MAX_XML_BYTES


class _ExplicitExportError(RuntimeError):
    """A content-free explicit-export failure safe to map to localized UI."""

    def __init__(self, code: str):
        self.code = code
        super().__init__(code)


def _markdown_paste_block(markdown: str) -> str:
    """Return markdown padded so block parsers see a standalone table."""
    return "\n" + markdown.strip() + "\n"


class TabledownApp(rumps.App):
    ICON_NAME = "tablemark_menu_40.png"
    ICON_NAME_OFF = "tablemark_menu_40_off.png"
    ICON_NAME_CHECK = "tablemark_menu_40_check.png"

    def __init__(self):
        log("app starting")
        self._clear_stale_status_item_visibility()
        # Deliberately NOT persisted: the toggle is a session pause. Starting
        # enabled every launch avoids "turned it off weeks ago, now 'broken'".
        # (Policy: CLAUDE.md § 설정 영속성 정책 — fill_blanks/language persist.)
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
        self._flash_timer = None  # icon flash in progress (success feedback)
        # Serializes read/convert/write with explicit clipboard exports so a
        # slow watcher conversion can never overwrite a newer manual result.
        self._clipboard_operation_lock = threading.Lock()
        # Allow only one explicit XML export at a time so an older result can
        # never finish after a newer request. NSAppleScript is main-thread-only,
        # so the action first updates the busy UI and then defers the read to the
        # next main-run-loop turn rather than executing inside the menu callback.
        self._explicit_export_lock = threading.Lock()
        self._explicit_export_active = None

        self.toggle_item = rumps.MenuItem(
            t("menu.toggle", self.lang),
            callback=self.toggle,
            key="t",
        )
        self.toggle_item.state = 1 if self.enabled else 0
        self._show_cmd_ctrl_shortcut(self.toggle_item)
        self.copy_xml_item = rumps.MenuItem(
            t("menu.copy_xml", self.lang),
            callback=self.copy_as_xml,
            key="x",
        )
        self._show_cmd_ctrl_shortcut(self.copy_xml_item)
        self.copy_excel_formulas_item = rumps.MenuItem(
            t("menu.copy_excel_formulas", self.lang),
            callback=self.copy_selected_excel_formulas,
            key="e",
        )
        self._show_cmd_ctrl_shortcut(self.copy_excel_formulas_item)
        self.copy_markdown_item = rumps.MenuItem(
            t("menu.copy_markdown", self.lang),
            callback=self.copy_as_markdown,
        )
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

        self.login_item_supported = login_item.is_supported()
        if self.login_item_supported:
            self.login_item_menu = rumps.MenuItem(
                t("menu.login_item", self.lang),
                callback=self.toggle_login_item,
            )
            self.login_item_menu.state = 1 if login_item.is_enabled() else 0
        else:
            self.login_item_menu = None

        # Support action: export a scrubbed local log for a bug report. Grouped
        # with Help (support), not Settings — it's an action, not a preference.
        self.diagnostics_item = rumps.MenuItem(
            t("menu.diagnostics", self.lang),
            callback=self.share_diagnostics,
        )
        self.help_item = rumps.MenuItem(t("menu.help", self.lang), callback=self.show_help)
        self.quit_item = rumps.MenuItem(t("menu.quit", self.lang), callback=self.quit_app)

        self.settings_item = rumps.MenuItem(t("menu.settings", self.lang))
        settings_children = [self.fill_blanks_item, self.language_item]
        if self.login_item_menu is not None:
            settings_children.append(self.login_item_menu)
        self.settings_item.update(settings_children)

        # Order: actions, then preferences, then help/support/quit.
        menu_items = [
            self.toggle_item,
            self.copy_markdown_item,
            self.copy_xml_item,
            self.copy_excel_formulas_item,
            None,  # separator
            self.settings_item,
            None,  # separator
            self.diagnostics_item,
            self.help_item,
        ]
        menu_items.append(self.quit_item)
        self.menu = menu_items

        # --- Global hotkeys: ⌘⌃X -> XML, ⌘⌃T -> toggle, ⌘⌃E -> formulas ---
        # Strong-ref on self so the Carbon callback trampoline survives GC. One
        # handler dispatches all by id (see hotkey.py). Graceful: if a hotkey
        # fails to register, its menu item still works.
        self.hotkeys = GlobalHotkeys()
        self.hotkeys.add(KEY_X, CMD | CONTROL, self.copy_as_xml)
        self.hotkeys.add(KEY_T, CMD | CONTROL, self.toggle)
        self.hotkeys.add(
            KEY_E, CMD | CONTROL, self.copy_selected_excel_formulas
        )
        self.hotkeys.register()

        self._stop_watcher = threading.Event()
        self._last_change_count = clipboard_change_count()
        self._watcher_thread = threading.Thread(
            target=self._watch_clipboard,
            name="TabledownClipboardWatcher",
            daemon=True,
        )
        self._watcher_thread.start()
        log("clipboard watcher started")

        # First-run welcome: scheduled on the run loop (fires after run()
        # starts) so the alert never blocks startup; one-shot via the
        # welcome_shown flag.
        self._welcome_timer = None
        if not load_welcome_shown():
            self._welcome_timer = rumps.Timer(self._show_welcome_once, 1.5)
            self._welcome_timer.start()

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

    @staticmethod
    def _show_cmd_ctrl_shortcut(item):
        """Display "⌘⌃<key>" next to a menu item to advertise its global hotkey.

        rumps' ``key=`` gives a Command-only equivalent; we OR in Control so the
        text matches the Carbon hotkey (⌘⌃X / ⌘⌃T / ⌘⌃E). Purely cosmetic
        discoverability — for a status-bar app the key equivalent only fires
        while the menu is open; the real global trigger is the Carbon hotkey. So
        never let this crash startup.
        """
        try:
            from AppKit import NSCommandKeyMask, NSControlKeyMask
            item._menuitem.setKeyEquivalentModifierMask_(
                NSCommandKeyMask | NSControlKeyMask
            )
        except Exception as exc:  # noqa: BLE001
            log(f"failed to set shortcut modifier mask: {exc}")

    # --- Localization ---

    def _language_option_title(self, code: str) -> str:
        return t(f"menu.language.{code}", self.lang)

    def _apply_language(self):
        """Refresh every menu title for the current language."""
        self.toggle_item.title = t("menu.toggle", self.lang)
        self._update_explicit_export_menu()
        self.fill_blanks_item.title = t("menu.fill_blanks", self.lang)
        self.fill_blanks_item._menuitem.setToolTip_(t("menu.fill_blanks_tooltip", self.lang))
        self.settings_item.title = t("menu.settings", self.lang)
        self.language_item.title = t("menu.language", self.lang)
        for code, item in self.language_options.items():
            item.title = self._language_option_title(code)
        if self.login_item_menu is not None:
            self.login_item_menu.title = t("menu.login_item", self.lang)
        self.diagnostics_item.title = t("menu.diagnostics", self.lang)
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
        """Schedule a merge-aware XML export of the Excel selection.

        The established stable-snapshot reader and exact XML serializer remain
        unchanged. The busy menu state is applied before the main-thread-only
        NSAppleScript read starts, one export runs at a time, and a newer user
        clipboard copy cancels the final write.
        """
        self._start_explicit_export(_EXPORT_TABLE)

    def copy_selected_excel_formulas(self, _sender):
        """Schedule a values/formulas/references XML export."""
        self._start_explicit_export(_EXPORT_FORMULAS)

    def copy_as_markdown(self, _sender):
        """Copy the selected Excel table with its displayed values and layout."""
        self._start_explicit_export(_EXPORT_MARKDOWN)

    def _start_explicit_export(self, export_kind: str) -> None:
        """Schedule one explicit copy action, or explain that one is already active."""
        if not self._explicit_export_lock.acquire(blocking=False):
            self._safe_alert(
                t("export.error_title", self.lang),
                t("export.error.in_progress", self.lang),
            )
            return

        try:
            # Let an in-flight watcher finish first, then establish the exact
            # clipboard generation this explicit action is allowed to replace.
            with self._clipboard_operation_lock:
                # Publish the watcher-pause signal before releasing this lock.
                # Otherwise an automatic rewrite can increment changeCount
                # after the snapshot and falsely look like a new user copy.
                self._explicit_export_active = export_kind
                expected_change_count = clipboard_change_count()
            fill_blanks = self.fill_blanks
            self._set_explicit_export_busy(export_kind)
            AppHelper.callAfter(
                self._run_explicit_export,
                export_kind,
                expected_change_count,
                fill_blanks,
            )
        except Exception as exc:  # noqa: BLE001 - scheduling failures are UI-safe
            self._explicit_export_lock.release()
            self._set_explicit_export_busy(None)
            log(f"start explicit XML export failed: {type(exc).__name__}")
            self._alert_explicit_export_error(export_kind, "execution_failed")

    def _run_explicit_export(
        self,
        export_kind: str,
        expected_change_count: int,
        fill_blanks: bool,
    ) -> None:
        """Run the explicit export on the main thread where NSAppleScript is valid."""
        outcome = "error"
        error_code = "execution_failed"
        try:
            outcome, error_code = self._perform_explicit_export(
                export_kind, expected_change_count, fill_blanks
            )
        except Exception as exc:  # noqa: BLE001 - gate/UI must always recover
            log(f"explicit XML export failed: {type(exc).__name__}")
        self._finish_explicit_export(export_kind, outcome, error_code)

    def _perform_explicit_export(
        self,
        export_kind: str,
        expected_change_count: int,
        fill_blanks: bool,
    ) -> tuple[str, str | None]:
        """Read Excel and write the requested table format, returning a content-free result."""
        outcome = "success"
        error_code = None
        notice_key = None
        output_html = None
        output_name = "Markdown" if export_kind == _EXPORT_MARKDOWN else "XML"
        stop_event = getattr(self, "_stop_watcher", None)
        try:
            if stop_event is not None and stop_event.is_set():
                raise _ExplicitExportError("cancelled")
            if export_kind == _EXPORT_MARKDOWN:
                selection = read_stable_selected_excel_table()
                table_html = excel_table_selection_to_html(selection)
                xml = _markdown_paste_block(html_table_to_markdown(
                    table_html,
                    fill_merged_headers=fill_blanks,
                    preserve_layout=True,
                ))
                output_html = (
                    '<html><head><meta charset="utf-8">'
                    '<style>td{white-space:pre-wrap}</style></head><body>'
                    + table_html + '</body></html>'
                )
                output_bytes = len(xml.encode("utf-8")) + len(output_html.encode("utf-8"))
                if output_bytes > MAX_MARKDOWN_EXPORT_BYTES:
                    raise _ExplicitExportError("output_too_large")
            elif export_kind == _EXPORT_TABLE:
                selection = read_stable_selected_excel_table()
                (
                    header_levels,
                    data_rows,
                    data_source_addresses,
                ) = excel_table_selection_to_model_with_sources(selection)
                if not data_rows:
                    raise _ExplicitExportError("no_table")
                filled_cell_addresses: tuple[str, ...] = ()
                if fill_blanks:
                    original_rows = [list(row) for row in data_rows]
                    fill_input = [
                        list(header_levels[-1]),
                        *[list(row) for row in data_rows],
                    ]
                    data_rows = forward_fill_key_columns(fill_input)[1:]
                    filled_cell_addresses = tuple(
                        dict.fromkeys(
                            data_source_addresses[row_index][column_index]
                            for row_index, (before, after) in enumerate(
                                zip(original_rows, data_rows, strict=True)
                            )
                            for column_index, (old_value, new_value) in enumerate(
                                zip(before, after, strict=True)
                            )
                            if old_value != new_value
                        )
                    )
                metadata = excel_table_selection_xml_metadata(
                    selection,
                    header_levels,
                    blank_fill_enabled=fill_blanks,
                    blank_fill_cells=filled_cell_addresses,
                )
                xml = model_to_xml(
                    header_levels,
                    data_rows,
                    metadata=metadata,
                )
            else:
                selection = read_stable_selected_excel_formulas()
                xml = formula_selection_to_ai_xml(selection)
                notice_key = formula_copy_notice_key(selection)

            if stop_event is not None and stop_event.is_set():
                raise _ExplicitExportError("cancelled")

            with self._clipboard_operation_lock:
                if stop_event is not None and stop_event.is_set():
                    raise _ExplicitExportError("cancelled")
                if output_html is not None:
                    write_table_clipboard(
                        xml, output_html, mark_generated=True,
                        expected_change_count=expected_change_count,
                    )
                else:
                    write_text_only_clipboard(
                        xml,
                        mark_generated=True,
                        expected_change_count=expected_change_count,
                    )

            # A notice describes successfully copied, incomplete context. Never
            # show it before the writer has verified the clipboard contents.
            if notice_key is not None:
                outcome, error_code = "notice", notice_key
            if export_kind != _EXPORT_FORMULAS:
                log(f"copied selected Excel table as {output_name}")
            else:
                log("copied selected Excel table with formulas as XML")
        except ClipboardChangedError:
            outcome = "error"
            error_code = "clipboard_changed"
            if export_kind != _EXPORT_FORMULAS:
                log(f"copy selected Excel table as {output_name} failed: clipboard_changed")
            else:
                log("copy Excel table with formulas failed: clipboard_changed")
        except ClipboardWriteError:
            outcome = "error"
            error_code = "clipboard_write_failed"
            if export_kind != _EXPORT_FORMULAS:
                log(
                    f"copy selected Excel table as {output_name} failed: "
                    "clipboard_write_failed"
                )
            else:
                log(
                    "copy Excel table with formulas failed: "
                    "clipboard_write_failed"
                )
        except (TableXmlTooLargeError, FormulaXmlTooLargeError):
            outcome = "error"
            error_code = "output_too_large"
            if export_kind != _EXPORT_FORMULAS:
                log(f"copy selected Excel table as {output_name} failed: output_too_large")
            else:
                log("copy Excel table with formulas failed: output_too_large")
        except ExcelFormulaError as exc:
            outcome = "error"
            error_code = exc.code
            if export_kind != _EXPORT_FORMULAS:
                log(f"copy selected Excel table as {output_name} failed: {exc.code}")
            else:
                log(f"copy Excel table with formulas failed: {exc.code}")
        except _ExplicitExportError as exc:
            outcome = "cancelled" if exc.code == "cancelled" else "error"
            error_code = exc.code
            if exc.code != "cancelled":
                if export_kind != _EXPORT_FORMULAS:
                    log(f"copy selected Excel table as {output_name} failed: {exc.code}")
                else:
                    log(f"copy Excel table with formulas failed: {exc.code}")
        except Exception as exc:  # noqa: BLE001 - never log workbook content
            outcome = "error"
            error_code = "execution_failed"
            if export_kind != _EXPORT_FORMULAS:
                log(
                    f"copy selected Excel table as {output_name} failed: "
                    f"{type(exc).__name__}"
                )
            else:
                log("copy Excel table with formulas failed: execution_failed")
        return outcome, error_code

    def _finish_explicit_export(
        self, export_kind: str, outcome: str, error_code: str | None
    ) -> None:
        """Restore menu state and report the verified export on the main thread."""
        try:
            stop_event = getattr(self, "_stop_watcher", None)
            if stop_event is not None and stop_event.is_set():
                return
            self._set_explicit_export_busy(None)
            if outcome in {"success", "notice"}:
                # A notification would request a new permission.  Keep the
                # existing permission-free one-second menu icon confirmation.
                self._flash_icon_success()
                if outcome == "notice" and error_code is not None:
                    self._safe_alert(
                        t("formula.copy_notice_title", self.lang),
                        t(f"formula.copy_notice.{error_code}", self.lang),
                    )
            elif outcome == "error" and error_code is not None:
                self._alert_explicit_export_error(export_kind, error_code)
        finally:
            self._explicit_export_lock.release()

    def _alert_explicit_export_error(self, export_kind: str, code: str) -> None:
        prefix = "formula" if export_kind == _EXPORT_FORMULAS else "table"
        key = f"{prefix}.error.{code}"
        if export_kind == _EXPORT_MARKDOWN:
            markdown_key = f"markdown.error.{code}"
            if t(markdown_key, self.lang) != markdown_key:
                key = markdown_key
        message = t(key, self.lang)
        if message == key:
            message = t(f"{prefix}.error.execution_failed", self.lang)
        self._safe_alert(
            t(f"{prefix}.error_title", self.lang),
            message,
        )

    def _set_explicit_export_busy(self, export_kind: str | None) -> None:
        """Expose progress in both the closed status item and its open menu."""
        self._explicit_export_active = export_kind
        # The menu closes before the main-thread Excel read begins, so a changed
        # menu item alone is not visible during a long selection. Keep a small
        # status-bar ellipsis beside the icon until success/failure restores it.
        try:
            self.title = "…" if export_kind is not None else None
        except Exception as exc:  # noqa: BLE001 - progress text is best effort
            log(f"explicit XML status title update failed: {type(exc).__name__}")
        self._update_explicit_export_menu()
        enabled = export_kind is None
        for item in (
            self.copy_markdown_item,
            self.copy_xml_item,
            self.copy_excel_formulas_item,
        ):
            try:
                item._menuitem.setEnabled_(enabled)
            except Exception as exc:  # noqa: BLE001 - state text is best effort
                log(f"explicit XML menu state update failed: {type(exc).__name__}")

    def _update_explicit_export_menu(self) -> None:
        """Apply localized normal/busy copy labels without losing active state."""
        active = getattr(self, "_explicit_export_active", None)
        self.copy_xml_item.title = t(
            "menu.copy_xml_busy" if active == _EXPORT_TABLE else "menu.copy_xml",
            self.lang,
        )
        self.copy_excel_formulas_item.title = t(
            (
                "menu.copy_excel_formulas_busy"
                if active == _EXPORT_FORMULAS
                else "menu.copy_excel_formulas"
            ),
            self.lang,
        )
        self.copy_markdown_item.title = t(
            "menu.copy_markdown_busy" if active == _EXPORT_MARKDOWN else "menu.copy_markdown",
            self.lang,
        )
        for item, key in (
            (self.copy_markdown_item, "menu.copy_markdown_tooltip"),
            (self.copy_xml_item, "menu.copy_xml_tooltip"),
            (self.copy_excel_formulas_item, "menu.copy_excel_formulas_tooltip"),
        ):
            item._menuitem.setToolTip_(t(key, self.lang))

    def _safe_alert(self, title: str, message: str):
        """rumps.alert that never lets a UI failure propagate into the run loop."""
        try:
            rumps.alert(title=title, message=message)
        except Exception as exc:  # noqa: BLE001
            log(f"alert failed: {exc}")

    # --- Success flash ---

    def _flash_icon_success(self, duration: float = 1.0):
        """Swap the menu bar icon to a checkmark briefly, then restore it.

        ``duration`` defaults to 1.0s for the deliberate XML action; the
        background watcher passes a shorter value so the every-copy feedback
        stays unobtrusive. MUST run on the main thread (the watcher hops here
        via AppHelper.callAfter) — touching ``self.icon`` or creating a
        rumps.Timer off the main run loop is not safe.
        """
        path = self._icon_path(self.ICON_NAME_CHECK)
        if path is None:
            return
        if self._flash_timer is not None:
            self._flash_timer.stop()
            self._flash_timer = None
        self.icon = path
        # rumps.Timer runs on the main run loop — safe to touch the icon.
        self._flash_timer = rumps.Timer(self._end_icon_flash, duration)
        self._flash_timer.start()

    def _end_icon_flash(self, timer):
        timer.stop()
        self._flash_timer = None
        self.icon = self._current_icon_path()

    @staticmethod
    def _clipboard_table_model(content, fill_blanks=False):
        """Extract a ``(header_levels, data_rows)`` model from text/HTML content.

        The explicit Excel-selection action supplies synthesized merge-aware
        HTML.  The XML and Markdown branches remain as conservative converter
        helpers and regression fixtures; the action deliberately does not fall
        back to stale clipboard content. ``header_levels`` is one list per
        header level so multi-level group headers survive. When ``fill_blanks``
        is set, blank cells in the left grouping columns are forward-filled (see
        forward_fill_key_columns) — the user-controlled option, off by default.
        """
        html = content.get("html", "")
        text = content.get("text", "")
        model = None
        if html and "<table" in html.lower():
            try:
                # Merge-aware: forward-fills rowspan, skips a full-width title
                # row, and keeps multi-level headers — see html_table_to_model.
                model = html_table_to_model(html)
            except MultipleTablesError:
                raise
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
        """Help alert: usage text, the running version, and a GitHub button.

        The version in the title is the only place a user can read it (there is
        no About window) — needed for bug reports. The legacy NSAlert API rumps
        wraps returns -1 (NSAlertOtherReturn) for the third button.
        """
        try:
            response = rumps.alert(
                title=f"Tabledown v{__version__}",
                message=t("help.message", self.lang),
                other=t("help.open_github", self.lang),
            )
        except Exception as exc:  # noqa: BLE001
            log(f"help alert failed: {exc}")
            return
        if response == -1:
            self._open_github()

    def share_diagnostics(self, _sender):
        """Write a scrubbed local log and reveal it in Finder for a bug report.

        Local only — nothing is sent anywhere. Touches no clipboard, so it can't
        race the watcher. On export failure we do NOT fall back to revealing the
        raw Tabledown.log: that file is unscrubbed and the user would attach it
        believing it was sanitized.
        """
        path = diagnostics.export_diagnostics()
        if path is None:
            self._safe_alert(
                t("table.error_title", self.lang),
                t("diagnostics.export_failed", self.lang),
            )
            return
        diagnostics.reveal(path)

    @staticmethod
    def _open_url(url):
        """Open a URL in the default browser (sandbox-safe via NSWorkspace, no subprocess)."""
        if not url:
            return
        try:
            from AppKit import NSWorkspace
            from Foundation import NSURL
            NSWorkspace.sharedWorkspace().openURL_(NSURL.URLWithString_(url))
        except Exception as exc:  # noqa: BLE001
            log(f"opening URL failed: {exc}")

    @classmethod
    def _open_github(cls):
        """Open the project page (sandbox-safe via NSWorkspace, no subprocess)."""
        cls._open_url(GITHUB_URL)

    def _show_welcome_once(self, timer):
        """First-run pointer: the app is menu-bar-only, so a fresh install
        otherwise looks like nothing happened. Marked shown before presenting
        so a failing alert can never loop the welcome on every launch."""
        timer.stop()
        save_welcome_shown(True)
        self._safe_alert(
            t("welcome.title", self.lang),
            t("welcome.intro", self.lang) + "\n\n" + t("help.message", self.lang),
        )

    def quit_app(self, _):
        self._stop_watcher.set()
        if getattr(self, "hotkeys", None) is not None:
            self.hotkeys.unregister()
        rumps.quit_application()

    # --- Clipboard watcher ---

    def _watch_clipboard(self):
        while not self._stop_watcher.is_set():
            time.sleep(0.1)
            if not self.enabled:
                continue
            if getattr(self, "_explicit_export_active", None) is not None:
                continue

            current_change_count = clipboard_change_count()
            if current_change_count == self._last_change_count:
                continue

            if self._augment_clipboard():
                self._last_change_count = current_change_count

    def _augment_clipboard(self):
        if getattr(self, "_explicit_export_active", None) is not None:
            return False
        try:
            with self._clipboard_operation_lock:
                # Recheck after acquiring the transaction lock to close the
                # race where an export starts after the outer watcher check.
                if getattr(self, "_explicit_export_active", None) is not None:
                    return False
                source_change_count = clipboard_change_count()
                content = read_clipboard()
                updated = self._converted_clipboard(content)
                if updated is None:
                    return True

                write_clipboard(
                    **updated,
                    mark_generated=True,
                    expected_change_count=source_change_count,
                )
            log("clipboard formats updated")
            # Make the otherwise-invisible background rewrite observable: a brief
            # icon checkmark so the user knows a table they copied was converted
            # (and can connect an unexpected paste to Tabledown, find the pause
            # toggle, etc.). Shorter than the manual XML flash since it fires on
            # every table copy. We run on the watcher thread, so hop to the main
            # thread before touching the icon. The mark_generated write above
            # makes the next watcher tick a no-op, so this never flash-loops.
            AppHelper.callAfter(self._flash_icon_success, 0.5)
            return True
        except ClipboardChangedError:
            # A newer external copy arrived while this generation was being
            # converted. Leave it untouched and let the next tick process it.
            return False
        except Exception as e:
            # Type only — this processes clipboard content and exc messages can
            # carry table data we must never write to the shareable log.
            log(f"clipboard update failed: {type(e).__name__}")
            return True

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

        # Same opt-in toggle as the XML path: when on, forward-fill the merged
        # header frame so a group-header band / left key column carry their value
        # instead of leaving blanks (value cells stay blank). getattr guards the
        # static-call test path (self is None) — defaults to off. See
        # _fill_header_frame.
        fill = bool(getattr(self, "fill_blanks", False))

        if has_html_table:
            # A table embedded in a document (headings, paragraphs, lists):
            # render its tables as Markdown in the text slot but KEEP the
            # original HTML slot. A Markdown editor reading text gets Markdown
            # tables; rich editors (Word, Excel) still read the original
            # <table> from HTML and paste a real table. Only RENDERED image
            # formats are dropped, never the HTML.
            if html_has_content_outside_table(html):
                converted = convert_document_tables(html, fill)
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
            markdown = html_table_to_markdown(html, fill)
            if text.strip() == markdown.strip():
                return None
            log("detected html table")
            return {
                "text": _markdown_paste_block(markdown),
                "drop_types": RENDERED_TABLE_TYPES,
            }

        return None

def main():
    # Capture failures that otherwise vanish (watcher-thread exceptions, native
    # faults) into the local log — purely local, nothing is transmitted.
    diagnostics.install_crash_hooks()
    TabledownApp().run()


if __name__ == "__main__":
    main()
