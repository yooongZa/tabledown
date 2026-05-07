"""Global hotkey listener for Cmd+Ctrl+M on macOS."""
import threading

from Quartz import (
    CFRunLoopAddSource,
    CFRunLoopGetCurrent,
    CFRunLoopRun,
    CFRunLoopStop,
    CFMachPortCreateRunLoopSource,
    CGEventGetFlags,
    CGEventGetIntegerValueField,
    CGEventMaskBit,
    CGPreflightListenEventAccess,
    CGRequestListenEventAccess,
    CGEventTapCreate,
    CGEventTapEnable,
    kCFRunLoopCommonModes,
    kCGEventFlagMaskCommand,
    kCGEventFlagMaskControl,
    kCGEventKeyDown,
    kCGEventTapDisabledByTimeout,
    kCGEventTapDisabledByUserInput,
    kCGEventTapOptionListenOnly,
    kCGHeadInsertEventTap,
    kCGKeyboardEventAutorepeat,
    kCGKeyboardEventKeycode,
    kCGSessionEventTap,
)

from .logger import log


class HotkeyListener:
    """Listens for Cmd+Ctrl+M system-wide."""

    HOTKEY_LABEL = "Cmd+Ctrl+M"
    KEY_CODE_M = 46
    REQUIRED_FLAGS = (
        kCGEventFlagMaskCommand
        | kCGEventFlagMaskControl
    )

    def __init__(self, callback):
        self.callback = callback
        self.event_tap = None
        self.run_loop = None
        self.thread = None
        self._callback_ref = None
        self._ready = threading.Event()
        self._startup_error = None

    def start(self):
        log(f"starting hotkey listener: {self.HOTKEY_LABEL}")
        self.thread = threading.Thread(
            target=self._run,
            name="TabledownHotkeyListener",
            daemon=True,
        )
        self.thread.start()
        self._ready.wait(timeout=1)
        if self._startup_error:
            raise self._startup_error

    def _run(self):
        self.run_loop = CFRunLoopGetCurrent()
        self._callback_ref = self._handle_event
        try:
            has_access = CGPreflightListenEventAccess()
            log(f"listen event access preflight: {has_access}")
            if not has_access:
                requested = CGRequestListenEventAccess()
                log(f"listen event access requested: {requested}")
        except Exception as e:
            log(f"listen event access check failed: {e}")

        self.event_tap = CGEventTapCreate(
            kCGSessionEventTap,
            kCGHeadInsertEventTap,
            kCGEventTapOptionListenOnly,
            CGEventMaskBit(kCGEventKeyDown),
            self._callback_ref,
            None,
        )

        if not self.event_tap:
            self._startup_error = RuntimeError(
                "전역 단축키 권한이 필요합니다. Accessibility/Input Monitoring을 확인하세요."
            )
            log("hotkey listener failed: CGEventTapCreate returned None")
            self._ready.set()
            return

        source = CFMachPortCreateRunLoopSource(None, self.event_tap, 0)
        CFRunLoopAddSource(self.run_loop, source, kCFRunLoopCommonModes)
        CGEventTapEnable(self.event_tap, True)
        log("hotkey listener started")
        self._ready.set()
        CFRunLoopRun()

    def _handle_event(self, proxy, event_type, event, _):
        if event_type in (kCGEventTapDisabledByTimeout, kCGEventTapDisabledByUserInput):
            CGEventTapEnable(self.event_tap, True)
            return event

        if event_type != kCGEventKeyDown:
            return event

        key_code = CGEventGetIntegerValueField(event, kCGKeyboardEventKeycode)
        is_repeat = CGEventGetIntegerValueField(event, kCGKeyboardEventAutorepeat)
        flags = CGEventGetFlags(event)

        if (
            key_code == self.KEY_CODE_M
            and not is_repeat
            and (flags & self.REQUIRED_FLAGS) == self.REQUIRED_FLAGS
        ):
            log("hotkey activated")
            self.callback()

        return event

    def stop(self):
        if self.event_tap:
            CGEventTapEnable(self.event_tap, False)
            self.event_tap = None
        if self.run_loop:
            CFRunLoopStop(self.run_loop)
            self.run_loop = None
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=0.5)
        self.thread = None
