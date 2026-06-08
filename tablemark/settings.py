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

# XML Pro subscription state (see store.py / the monetization plan §3.4). These
# are a *cache* of the StoreKit entitlement, written by the transaction observer
# on purchase/renew/restore. pro_active is the gating bool; pro_expires is the
# epoch timestamp used for the MVP "now < expiry (+grace)" check.
PRO_ACTIVE_KEY = "com.tabledown.app.pro_active"
PRO_EXPIRES_KEY = "com.tabledown.app.pro_expires"
PRO_ACTIVE_DEFAULT = False
PRO_EXPIRES_DEFAULT = 0.0


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


def load_pro_active() -> bool:
    """Read the cached XML Pro 'active' flag; default when unset."""
    try:
        defaults = NSUserDefaults.standardUserDefaults()
        if defaults.objectForKey_(PRO_ACTIVE_KEY) is None:
            return PRO_ACTIVE_DEFAULT
        return bool(defaults.boolForKey_(PRO_ACTIVE_KEY))
    except Exception as exc:  # noqa: BLE001 - never let a settings read crash startup
        log(f"failed to read pro_active: {exc}")
        return PRO_ACTIVE_DEFAULT


def save_pro_active(active: bool) -> None:
    """Persist the cached XML Pro 'active' flag."""
    try:
        NSUserDefaults.standardUserDefaults().setBool_forKey_(bool(active), PRO_ACTIVE_KEY)
    except Exception as exc:  # noqa: BLE001
        log(f"failed to save pro_active: {exc}")


def load_pro_expires() -> float:
    """Read the cached XML Pro expiry (epoch seconds); default when unset."""
    try:
        defaults = NSUserDefaults.standardUserDefaults()
        if defaults.objectForKey_(PRO_EXPIRES_KEY) is None:
            return PRO_EXPIRES_DEFAULT
        return float(defaults.doubleForKey_(PRO_EXPIRES_KEY))
    except Exception as exc:  # noqa: BLE001 - never let a settings read crash startup
        log(f"failed to read pro_expires: {exc}")
        return PRO_EXPIRES_DEFAULT


def save_pro_expires(expires: float) -> None:
    """Persist the cached XML Pro expiry (epoch seconds)."""
    try:
        NSUserDefaults.standardUserDefaults().setDouble_forKey_(float(expires), PRO_EXPIRES_KEY)
    except Exception as exc:  # noqa: BLE001
        log(f"failed to save pro_expires: {exc}")
