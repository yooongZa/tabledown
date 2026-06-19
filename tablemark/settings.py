"""Persisted user settings backed by NSUserDefaults.

Currently just the blank-fill toggle. Kept out of i18n.py (which owns the
language preference) so each concern has one home, mirroring login_item.py.
"""
from __future__ import annotations

from Foundation import NSUserDefaults

from .logger import log

FILL_BLANKS_KEY = "com.tabledown.app.fill_blanks"
# Off by default (the "safe" choice): forward-filling blank cells can invent
# data that wasn't there — a blank may be "genuinely empty", not "same as
# above" — so it is strictly opt-in. Matches what data tools (pandas read_html,
# Power Query) do: never auto-fill, the user asks for it. See invariant 5 in
# CLAUDE.md.
FILL_BLANKS_DEFAULT = False

# First-run welcome: the app is menu-bar-only (LSUIElement, no Dock icon), so
# without a one-time pointer a fresh install looks like "nothing happened".
WELCOME_SHOWN_KEY = "com.tabledown.app.welcome_shown"


def load_fill_blanks() -> bool:
    """Read the blank-fill toggle; fall back to the default when unset."""
    try:
        defaults = NSUserDefaults.standardUserDefaults()
        if defaults.objectForKey_(FILL_BLANKS_KEY) is None:
            return FILL_BLANKS_DEFAULT
        return bool(defaults.boolForKey_(FILL_BLANKS_KEY))
    except Exception as exc:  # noqa: BLE001 - never let a settings read crash startup
        log(f"failed to read fill_blanks: {exc}")
        return FILL_BLANKS_DEFAULT


def save_fill_blanks(enabled: bool) -> None:
    """Persist the blank-fill toggle."""
    try:
        NSUserDefaults.standardUserDefaults().setBool_forKey_(bool(enabled), FILL_BLANKS_KEY)
    except Exception as exc:  # noqa: BLE001
        log(f"failed to save fill_blanks: {exc}")


def load_welcome_shown() -> bool:
    """Read whether the first-run welcome has been shown (False when unset)."""
    try:
        return bool(NSUserDefaults.standardUserDefaults().boolForKey_(WELCOME_SHOWN_KEY))
    except Exception as exc:  # noqa: BLE001 - never let a settings read crash startup
        log(f"failed to read welcome_shown: {exc}")
        return True  # fail closed: better to skip the welcome than loop it


def save_welcome_shown(shown: bool = True) -> None:
    """Persist that the first-run welcome has been shown."""
    try:
        NSUserDefaults.standardUserDefaults().setBool_forKey_(bool(shown), WELCOME_SHOWN_KEY)
    except Exception as exc:  # noqa: BLE001
        log(f"failed to save welcome_shown: {exc}")
