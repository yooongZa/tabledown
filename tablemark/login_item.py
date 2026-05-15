"""SMAppService wrapper for the "Open at Login" toggle.

macOS 13+ exposes `SMAppService.mainAppService` as the modern replacement for the
deprecated `SMLoginItemSetEnabled` and bare LaunchAgent plists. It works under
App Sandbox and Mac App Store distribution and persists state in the system
"Login Items" pane — no NSUserDefaults shadowing.

The wrapper degrades gracefully: when ServiceManagement is unavailable or macOS
is too old, `is_supported()` returns False and the menu can hide the toggle.
"""
from __future__ import annotations

import platform

from .logger import log

try:
    import ServiceManagement as _SM
except ImportError:
    _SM = None

_SMAppService = getattr(_SM, "SMAppService", None) if _SM else None
_STATUS_ENABLED = getattr(_SM, "SMAppServiceStatusEnabled", 1) if _SM else 1


def _macos_at_least_13() -> bool:
    raw = platform.mac_ver()[0]
    if not raw:
        return False
    try:
        major = int(raw.split(".")[0])
    except ValueError:
        return False
    return major >= 13


def is_supported() -> bool:
    """True when SMAppService.mainAppService can be used on this machine."""
    return _SMAppService is not None and _macos_at_least_13()


def _service():
    if not is_supported():
        return None
    try:
        return _SMAppService.mainAppService()
    except Exception as exc:
        log(f"SMAppService.mainAppService failed: {exc}")
        return None


def is_enabled() -> bool:
    """True when the app is currently registered to launch at login."""
    svc = _service()
    if svc is None:
        return False
    try:
        return int(svc.status()) == int(_STATUS_ENABLED)
    except Exception as exc:
        log(f"SMAppService status check failed: {exc}")
        return False


def set_enabled(enabled: bool) -> bool:
    """Register or unregister the app as a login item. Returns the new state."""
    svc = _service()
    if svc is None:
        return False
    try:
        if enabled:
            ok, err = svc.registerAndReturnError_(None)
        else:
            ok, err = svc.unregisterAndReturnError_(None)
        if not ok:
            log(f"SMAppService {'register' if enabled else 'unregister'} failed: {err}")
    except Exception as exc:
        log(f"SMAppService toggle raised: {exc}")
    return is_enabled()
