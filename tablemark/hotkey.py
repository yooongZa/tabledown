"""Carbon global hotkey (⌘⌃C) — permission-free, App Sandbox compatible.

Registers a system-wide hotkey via Carbon's ``RegisterEventHotKey`` +
``InstallEventHandler``. This path needs **no** Accessibility / Input Monitoring
permission (unlike a CGEventTap or a pynput keyboard hook), which preserves
Tabledown's "권한 0개" (zero-permissions) selling point. It is also the standard
API many sandboxed Mac App Store apps use for global shortcuts.

Why ctypes and not a PyObjC framework: the modern PyObjC distribution ships no
``Carbon``/``HIToolbox`` module (pyobjc-framework-Carbon is legacy), but the
symbols live in the always-present system ``Carbon.framework``. We bind the four
C functions we need directly with ctypes. The installed event handler fires on
the main run loop that rumps' NSApplication already runs, so the Python callback
executes on the main thread — safe to touch UI / clipboard from it.

Failure is graceful: if anything fails to load or register we log and carry on;
``registered`` stays False and the corresponding menu item keeps working.
"""
from __future__ import annotations

import ctypes
from ctypes import (
    CFUNCTYPE,
    POINTER,
    Structure,
    byref,
    c_int32,
    c_uint32,
    c_void_p,
)

from .logger import log

# ⌘⌃C: Command + Control + C. Carbon modifier masks (Events.h) and the ANSI 'C'
# virtual keycode (kVK_ANSI_C == 8). These are the values the spec fixes.
_CMD_KEY = 0x0100  # cmdKey
_CONTROL_KEY = 0x1000  # controlKey
_KEY_C = 8  # kVK_ANSI_C

# Carbon event constants.
_EVENT_CLASS_KEYBOARD = (
    (ord("k") << 24) | (ord("e") << 16) | (ord("y") << 8) | ord("b")
)  # 'keyb'
_EVENT_HOTKEY_PRESSED = 6  # kEventHotKeyPressed
_HOTKEY_SIGNATURE = (
    (ord("T") << 24) | (ord("D") << 16) | (ord("W") << 8) | ord("N")
)  # 'TDWN'
_HOTKEY_ID = 1
_NOERR = 0


class _EventTypeSpec(Structure):
    _fields_ = [("eventClass", c_uint32), ("eventKind", c_uint32)]


class _EventHotKeyID(Structure):
    _fields_ = [("signature", c_uint32), ("id", c_uint32)]


# Obj-C/C callback signature: OSStatus (*)(EventHandlerCallRef, EventRef, void*).
_HANDLER_PROTO = CFUNCTYPE(c_int32, c_void_p, c_void_p, c_void_p)


def _load_carbon():
    """Load Carbon.framework and configure the C signatures we use, or None."""
    try:
        carbon = ctypes.CDLL("/System/Library/Frameworks/Carbon.framework/Carbon")
    except OSError as exc:
        log(f"Carbon framework load failed: {exc}")
        return None

    try:
        carbon.GetApplicationEventTarget.restype = c_void_p

        carbon.InstallEventHandler.argtypes = [
            c_void_p,  # EventTargetRef
            c_void_p,  # EventHandlerUPP (our CFUNCTYPE cast to void*)
            c_uint32,  # numTypes
            POINTER(_EventTypeSpec),
            c_void_p,  # userData
            POINTER(c_void_p),  # outRef
        ]
        carbon.InstallEventHandler.restype = c_int32

        carbon.RegisterEventHotKey.argtypes = [
            c_uint32,  # hotKeyCode
            c_uint32,  # modifiers
            _EventHotKeyID,
            c_void_p,  # EventTargetRef
            c_uint32,  # options
            POINTER(c_void_p),  # outRef
        ]
        carbon.RegisterEventHotKey.restype = c_int32

        carbon.UnregisterEventHotKey.argtypes = [c_void_p]
        carbon.UnregisterEventHotKey.restype = c_int32
    except AttributeError as exc:
        log(f"Carbon missing expected symbol: {exc}")
        return None

    return carbon


class GlobalHotkey:
    """Registers ⌘⌃C and invokes a Python callback on press.

    Usage:
        hk = GlobalHotkey(self.copy_as_xml)
        hk.register()
        ...
        hk.unregister()   # optional; on quit

    ``register()`` returns True on success. On failure it logs and returns False;
    callers should keep their menu item functional regardless (the hotkey is an
    accelerator, not the only entry point).
    """

    def __init__(self, callback):
        self._callback = callback
        self._carbon = None
        self._handler_ref = c_void_p()
        self._hotkey_ref = c_void_p()
        # Strong refs: the CFUNCTYPE trampoline and its void* cast must outlive
        # registration or Carbon would call freed memory when the key is pressed.
        self._handler_func = None
        self._handler_func_ptr = None
        self.registered = False

    def register(self) -> bool:
        """Install the event handler and register ⌘⌃C. True on success."""
        if self.registered:
            return True
        carbon = _load_carbon()
        if carbon is None:
            return False
        self._carbon = carbon
        try:
            target = carbon.GetApplicationEventTarget()
            if not target:
                log("GetApplicationEventTarget returned null")
                return False

            self._handler_func = _HANDLER_PROTO(self._on_hotkey)
            self._handler_func_ptr = ctypes.cast(self._handler_func, c_void_p)
            spec = _EventTypeSpec(_EVENT_CLASS_KEYBOARD, _EVENT_HOTKEY_PRESSED)
            status = carbon.InstallEventHandler(
                target,
                self._handler_func_ptr,
                1,
                byref(spec),
                None,
                byref(self._handler_ref),
            )
            if status != _NOERR:
                log(f"InstallEventHandler failed: status={status}")
                return False

            hotkey_id = _EventHotKeyID(_HOTKEY_SIGNATURE, _HOTKEY_ID)
            status = carbon.RegisterEventHotKey(
                _KEY_C,
                _CMD_KEY | _CONTROL_KEY,
                hotkey_id,
                target,
                0,
                byref(self._hotkey_ref),
            )
            if status != _NOERR:
                log(f"RegisterEventHotKey failed: status={status}")
                return False

            self.registered = True
            log("global hotkey ⌘⌃C registered")
            return True
        except Exception as exc:  # noqa: BLE001 - never let hotkey setup crash startup
            log(f"hotkey registration raised: {exc}")
            return False

    def _on_hotkey(self, _call_ref, _event, _user_data) -> int:
        """Carbon handler trampoline — runs on the main run loop thread."""
        try:
            # Carbon passes no sender; our callback (copy_as_xml) takes one arg.
            self._callback(None)
        except Exception as exc:  # noqa: BLE001 - a handler must always return
            log(f"hotkey callback raised: {exc}")
        return _NOERR

    def unregister(self) -> None:
        """Unregister the hotkey (best effort). Safe to call when not registered."""
        if self._carbon is not None and self._hotkey_ref:
            try:
                self._carbon.UnregisterEventHotKey(self._hotkey_ref)
            except Exception as exc:  # noqa: BLE001
                log(f"UnregisterEventHotKey raised: {exc}")
        self._hotkey_ref = c_void_p()
        self.registered = False
