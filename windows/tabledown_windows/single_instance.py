"""Single-instance guard for the Windows tray app.

macOS gets this for free — LaunchServices refuses to start a second copy of a
``.app`` bundle. On Windows nothing stops a second ``Tabledown.exe`` from
spawning: a manual launch on top of the MSIX StartupTask (launch-at-login), a
double double-click, or a stale copy a crash left behind. Each extra process
adds another tray icon *and* another clipboard watcher — and two watchers racing
on the clipboard can stomp each other's conversions (the invariants in CLAUDE.md
assume one writer). So we refuse to start a second instance.

The guard is a named mutex (``CreateMutexW``). The first process creates it; any
later process sees ``ERROR_ALREADY_EXISTS`` and bails out of :func:`main`. The
handle is held for the whole process lifetime (stashed in a module global, never
closed) so the mutex lives exactly as long as the running instance and vanishes
when it exits.

The name lives in the per-session ``Local\\`` namespace, so the limit is one
instance per interactive user session — fast-user-switching gets one each —
rather than a single ``Global\\`` instance for the whole machine.

Degrades safely: on a non-Windows host (the cross-platform test runners) the
kernel call is unavailable, so :func:`_create_mutex` reports "no sibling" and the
app is allowed to start. The kernel lookup is deferred into the function so
importing this module never fails off Windows.
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes

from .logger import log

# Per-session name: one tray instance per interactive user session. Keep this in
# sync with nothing else — it is a private kernel-object name, not user-visible.
MUTEX_NAME = "Local\\TabledownSingleInstance"

# CreateMutexW sets last-error to this when the named mutex already existed.
_ERROR_ALREADY_EXISTS = 183

# Held for the process lifetime so the mutex stays alive; never closed. A second
# instance closing/releasing would let a third in, so the owner simply keeps it.
_held_handle = None


def _create_mutex(name: str):
    """Create or open a named mutex. Returns ``(handle, already_existed)``.

    Thin wrapper over ``CreateMutexW`` so the single-instance logic can be tested
    with a stubbed kernel call on non-Windows hosts. On a host without
    ``kernel32`` (Linux/macOS) returns ``(None, False)`` — i.e. "no sibling, go
    ahead" — since there is nothing to guard against there.
    """
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    except (AttributeError, OSError):
        # ctypes.WinDLL is defined only on Windows; absent it, no guard applies.
        return None, False
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
    handle = kernel32.CreateMutexW(None, False, name)
    # use_last_error=True routes the Win32 last-error into ctypes' per-thread
    # store, read here before any other ctypes call can clobber it.
    already_existed = ctypes.get_last_error() == _ERROR_ALREADY_EXISTS
    return handle, already_existed


def acquire_single_instance(name: str = MUTEX_NAME) -> bool:
    """Claim the single-instance slot.

    Returns ``True`` when this process is the first/only instance (the caller
    should proceed), ``False`` when a sibling instance already holds the slot
    (the caller should exit quietly). On the first successful claim the mutex
    handle is kept alive in a module global for the process lifetime so the mutex
    outlives this call.

    A failed mutex creation (NULL handle, e.g. an unexpected kernel error)
    degrades to ``True`` — better to run a possible duplicate than to refuse to
    start at all.
    """
    global _held_handle
    handle, already_existed = _create_mutex(name)
    if already_existed:
        log("another Tabledown instance is already running; exiting")
        return False
    _held_handle = handle  # keep alive for the process lifetime
    return True
