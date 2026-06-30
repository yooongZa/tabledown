"""Carbon global hotkeys — permission-free, App Sandbox compatible.

Registers system-wide hotkeys via Carbon's ``RegisterEventHotKey`` +
``InstallEventHandler``. This path needs **no** Accessibility / Input Monitoring
permission (unlike a CGEventTap or a pynput keyboard hook), which preserves
Tabledown's "권한 0개" (zero-permissions) selling point. It is also the standard
API many sandboxed Mac App Store apps use for global shortcuts.

Tabledown binds two hotkeys: ⌘⌃X → copy-as-XML and ⌘⌃T → toggle conversion.

⚠️ One handler, dispatch by id (회귀 금지). Carbon delivers EVERY
``kEventHotKeyPressed`` to the handlers installed on the application event
target. A handler that always returns ``noErr`` *consumes* the event, so
installing one such handler per hotkey is wrong: only the most-recently-installed
handler ever runs — for *both* hotkeys, firing the wrong callback. So we install
exactly ONE handler and read the fired hotkey's id from the event
(``GetEventParameter`` → ``EventHotKeyID``) to pick the right callback. Do not
"simplify" this into per-hotkey handlers.

Why ctypes and not a PyObjC framework: the modern PyObjC distribution ships no
``Carbon``/``HIToolbox`` module (pyobjc-framework-Carbon is legacy), but the
symbols live in the always-present system ``Carbon.framework``. We bind the C
functions we need directly with ctypes. The installed event handler fires on the
main run loop that rumps' NSApplication already runs, so the Python callbacks
execute on the main thread — safe to touch UI / clipboard from them.

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
    c_ulong,
    c_void_p,
)

from .logger import log

# Carbon modifier masks (Events.h). Public so callers compose the combo they
# want without re-deriving the magic numbers.
CMD = 0x0100  # cmdKey
CONTROL = 0x1000  # controlKey
SHIFT = 0x0200  # shiftKey
OPTION = 0x0800  # optionKey

# ANSI virtual keycodes (Events.h, kVK_ANSI_*). X == 7, T == 17.
KEY_X = 7  # kVK_ANSI_X (mnemonic for XML)
KEY_T = 17  # kVK_ANSI_T

# Carbon event constants.
_EVENT_CLASS_KEYBOARD = (
    (ord("k") << 24) | (ord("e") << 16) | (ord("y") << 8) | ord("b")
)  # 'keyb'
_EVENT_HOTKEY_PRESSED = 6  # kEventHotKeyPressed
_HOTKEY_SIGNATURE = (
    (ord("T") << 24) | (ord("D") << 16) | (ord("W") << 8) | ord("N")
)  # 'TDWN'
# GetEventParameter: read the EventHotKeyID carried by a hotkey event.
_PARAM_DIRECT_OBJECT = (
    (ord("-") << 24) | (ord("-") << 16) | (ord("-") << 8) | ord("-")
)  # kEventParamDirectObject '----'
_TYPE_EVENT_HOTKEY_ID = (
    (ord("h") << 24) | (ord("k") << 16) | (ord("i") << 8) | ord("d")
)  # typeEventHotKeyID 'hkid'
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

        # OSStatus GetEventParameter(EventRef, OSType name, OSType type,
        #   OSType* outActualType, ByteCount bufSize, ByteCount* outSize, void* out)
        # ByteCount is `unsigned long` (8 bytes on 64-bit macOS) -> c_ulong.
        carbon.GetEventParameter.argtypes = [
            c_void_p,  # EventRef
            c_uint32,  # inName (OSType)
            c_uint32,  # inDesiredType (OSType)
            POINTER(c_uint32),  # outActualType (NULL ok)
            c_ulong,  # inBufferSize (ByteCount)
            POINTER(c_ulong),  # outActualSize (NULL ok)
            c_void_p,  # outData
        ]
        carbon.GetEventParameter.restype = c_int32
    except AttributeError as exc:
        log(f"Carbon missing expected symbol: {exc}")
        return None

    return carbon


class GlobalHotkeys:
    """Registers one or more system-wide hotkeys behind a single Carbon handler.

    Usage:
        hk = GlobalHotkeys()
        hk.add(KEY_X, CMD | CONTROL, self.copy_as_xml)
        hk.add(KEY_T, CMD | CONTROL, self.toggle)
        hk.register()
        ...
        hk.unregister()   # optional; on quit

    Each ``callback`` is invoked with a single positional argument (``None``) so
    rumps menu callbacks (``def cb(self, _sender)``) work unchanged. ``register()``
    returns True if at least one hotkey was registered. Failures are logged and
    the corresponding menu items keep working — a hotkey is only an accelerator.
    """

    def __init__(self):
        self._carbon = None
        self._handler_ref = c_void_p()
        # Strong refs: the CFUNCTYPE trampoline and its void* cast must outlive
        # registration or Carbon would call freed memory when a key is pressed.
        self._handler_func = None
        self._handler_func_ptr = None
        self._specs = []  # queued (key_code, modifiers, callback) for register()
        self._bindings = {}  # hotkey id -> callback (the live registrations)
        self._hotkey_refs = {}  # hotkey id -> EventHotKeyRef (c_void_p)
        self._next_id = 1
        self.registered = False

    def add(self, key_code: int, modifiers: int, callback) -> "GlobalHotkeys":
        """Queue a hotkey to bind. Call before ``register()``. Chainable."""
        self._specs.append((key_code, modifiers, callback))
        return self

    def register(self) -> bool:
        """Install the shared handler and register every queued hotkey.

        Returns True when the handler installed and at least one hotkey bound.
        """
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

            for key_code, modifiers, callback in self._specs:
                hk_id = self._next_id
                self._next_id += 1
                ref = c_void_p()
                hotkey_id = _EventHotKeyID(_HOTKEY_SIGNATURE, hk_id)
                status = carbon.RegisterEventHotKey(
                    key_code,
                    modifiers,
                    hotkey_id,
                    target,
                    0,
                    byref(ref),
                )
                if status != _NOERR:
                    log(f"RegisterEventHotKey failed (code={key_code}): status={status}")
                    continue
                self._bindings[hk_id] = callback
                self._hotkey_refs[hk_id] = ref

            self.registered = bool(self._bindings)
            if self.registered:
                log(f"global hotkeys registered: {len(self._bindings)}")
            return self.registered
        except Exception as exc:  # noqa: BLE001 - never let hotkey setup crash startup
            log(f"hotkey registration raised: {exc}")
            return False

    def _fired_hotkey_id(self, event) -> int:
        """Read the fired hotkey's id from the event, or 0 if unavailable."""
        if not event or self._carbon is None:
            return 0
        hkid = _EventHotKeyID()
        status = self._carbon.GetEventParameter(
            event,
            _PARAM_DIRECT_OBJECT,
            _TYPE_EVENT_HOTKEY_ID,
            None,
            ctypes.sizeof(hkid),
            None,
            byref(hkid),
        )
        if status != _NOERR:
            return 0
        return int(hkid.id)

    def _on_hotkey(self, _call_ref, event, _user_data) -> int:
        """Carbon handler trampoline — runs on the main run loop thread."""
        try:
            callback = self._bindings.get(self._fired_hotkey_id(event))
            if callback is None and len(self._bindings) == 1:
                # Only one hotkey bound: even if the id read failed, the event
                # can only be ours, so fire it rather than drop a real press.
                callback = next(iter(self._bindings.values()))
            if callback is not None:
                callback(None)
        except Exception as exc:  # noqa: BLE001 - a handler must always return
            log(f"hotkey callback raised: {exc}")
        return _NOERR

    def unregister(self) -> None:
        """Unregister every hotkey (best effort). Safe when nothing is bound."""
        if self._carbon is not None:
            for ref in self._hotkey_refs.values():
                if ref:
                    try:
                        self._carbon.UnregisterEventHotKey(ref)
                    except Exception as exc:  # noqa: BLE001
                        log(f"UnregisterEventHotKey raised: {exc}")
        self._hotkey_refs.clear()
        self._bindings.clear()
        self.registered = False
