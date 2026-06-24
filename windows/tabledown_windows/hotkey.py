"""Windows global hotkey (Ctrl+Alt+T) via Win32 ``RegisterHotKey``.

macOS binds global hotkeys through Carbon (see ``tablemark/hotkey.py``); the
Windows analog is user32 ``RegisterHotKey``. A hotkey registered with a NULL
window handle posts ``WM_HOTKEY`` to the *registering thread's* message queue, so
this module runs a private thread that both registers the key and pumps messages.
pystray owns its own message loop on another thread and we must neither block it
nor share its queue — hence the dedicated thread.

No extra permission, pure ctypes (user32 + kernel32). Degrades safely off
Windows (the cross-platform test runner): ``_load_user32`` returns None and
``start()`` reports "not registered" instead of raising — the tray menu item
keeps working, so the hotkey is only an accelerator (parity with the macOS
graceful-failure contract).

The worker thread is a daemon, so a hard exit can't hang on it; ``stop()`` still
unregisters cleanly by posting ``WM_QUIT`` to break ``GetMessageW``.
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes
import threading

from .logger import log

# RegisterHotKey fsModifiers (winuser.h).
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000  # one WM_HOTKEY per press, not autorepeat-spammed

# Virtual-key code for 'T' (VK is the uppercase ASCII for letters/digits).
VK_T = 0x54

# Window messages.
_WM_QUIT = 0x0012
_WM_HOTKEY = 0x0312

# Our single hotkey's id. The id namespace is per-thread for a NULL-hwnd
# registration, so a fixed 1 is fine.
_HOTKEY_ID = 1


def _load_user32():
    """Return the user32 WinDLL with our signatures set, or None off Windows.

    Deferred like ``single_instance._create_mutex`` so importing this module
    never fails on the non-Windows test host, where there is no global hotkey to
    bind. ``use_last_error`` routes Win32 errors into ctypes' per-thread store.
    """
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
    except (AttributeError, OSError):
        # ctypes.WinDLL exists only on Windows; absent it, no hotkey applies.
        return None
    user32.RegisterHotKey.argtypes = (
        wintypes.HWND,
        ctypes.c_int,
        wintypes.UINT,
        wintypes.UINT,
    )
    user32.RegisterHotKey.restype = wintypes.BOOL
    user32.UnregisterHotKey.argtypes = (wintypes.HWND, ctypes.c_int)
    user32.UnregisterHotKey.restype = wintypes.BOOL
    user32.GetMessageW.argtypes = (
        ctypes.POINTER(wintypes.MSG),
        wintypes.HWND,
        wintypes.UINT,
        wintypes.UINT,
    )
    # GetMessage returns -1 on error, so read it signed (BOOL is a signed int).
    user32.GetMessageW.restype = ctypes.c_int
    user32.PostThreadMessageW.argtypes = (
        wintypes.DWORD,
        wintypes.UINT,
        wintypes.WPARAM,
        wintypes.LPARAM,
    )
    user32.PostThreadMessageW.restype = wintypes.BOOL
    return user32


def _current_thread_id() -> int:
    """The calling thread's id (for PostThreadMessageW), or 0 off Windows."""
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    except (AttributeError, OSError):
        return 0
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    return int(kernel32.GetCurrentThreadId())


class GlobalHotkey:
    """Registers one system-wide hotkey and runs a callback when it fires.

    Usage:
        hk = GlobalHotkey(MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, VK_T, app._toggle_from_hotkey)
        hk.start()
        ...
        hk.stop()   # on quit

    ``start()`` returns True when the hotkey is live. Failure (off Windows, or the
    combo already owned by another app) logs and returns False; the tray menu item
    keeps working, so the hotkey is a pure accelerator. ``callback`` is invoked
    with no arguments, on the hotkey thread.
    """

    def __init__(self, modifiers: int, vk: int, callback):
        self._modifiers = modifiers
        self._vk = vk
        self._callback = callback
        self._user32 = None
        self._thread = None
        self._thread_id = 0
        self._ready = threading.Event()
        self.registered = False

    def start(self) -> bool:
        """Spawn the worker thread, register the hotkey, and start pumping.

        Blocks briefly until the worker reports the registration outcome so the
        return value (and ``registered``) reflect reality.
        """
        if self._thread is not None:
            return self.registered
        user32 = _load_user32()
        if user32 is None:
            log("global hotkey unavailable (no user32)")
            return False
        self._user32 = user32
        self._ready.clear()
        self._thread = threading.Thread(
            target=self._run,
            args=(user32,),
            name="TabledownWindowsHotkey",
            daemon=True,
        )
        self._thread.start()
        self._ready.wait(timeout=2.0)
        return self.registered

    def _run(self, user32) -> None:
        # RegisterHotKey + GetMessageW must share a thread: a NULL-hwnd hotkey
        # posts WM_HOTKEY to *this* thread's queue.
        self._thread_id = _current_thread_id()
        self.registered = self._register(user32)
        self._ready.set()
        if not self.registered:
            return
        try:
            self._pump(user32)
        finally:
            self._unregister(user32)
            self.registered = False

    def _register(self, user32) -> bool:
        ok = bool(user32.RegisterHotKey(None, _HOTKEY_ID, self._modifiers, self._vk))
        if not ok:
            # Most often: another app already owns the combo. ctypes last-error
            # is available via get_last_error() but we only need pass/fail here.
            log("RegisterHotKey failed (combo busy or unsupported)")
        return ok

    def _unregister(self, user32) -> None:
        try:
            user32.UnregisterHotKey(None, _HOTKEY_ID)
        except Exception as exc:  # noqa: BLE001
            log(f"UnregisterHotKey raised: {type(exc).__name__}")

    def _pump(self, user32) -> None:
        msg = wintypes.MSG()
        while True:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret in (0, -1):  # 0 = WM_QUIT retrieved, -1 = error
                break
            self._handle_message(msg)

    def _handle_message(self, msg) -> None:
        if msg.message == _WM_HOTKEY and int(msg.wParam) == _HOTKEY_ID:
            try:
                self._callback()
            except Exception as exc:  # noqa: BLE001 - keep the pump alive
                log(f"hotkey callback raised: {type(exc).__name__}")

    def stop(self) -> None:
        """Unregister and stop pumping. Safe to call before/without start()."""
        if self._thread is None:
            return
        if self._user32 is not None and self._thread_id:
            try:
                # Wake GetMessageW so the loop exits and _unregister runs.
                self._user32.PostThreadMessageW(self._thread_id, _WM_QUIT, 0, 0)
            except Exception as exc:  # noqa: BLE001
                log(f"PostThreadMessageW raised: {type(exc).__name__}")
        self._thread.join(timeout=2.0)
        self._thread = None
        self._thread_id = 0
