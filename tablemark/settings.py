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
