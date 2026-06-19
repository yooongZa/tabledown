"""Local-only diagnostics for Tabledown (macOS).

This module NEVER sends anything off the machine. It only (1) captures failures
that currently vanish from the local log and (2) lets the user *manually* export
a scrubbed log to attach to a bug report. That keeps the shipped privacy promise
intact — "no network connections, no telemetry ... nothing is sent to any
external server" (PRIVACY.md). See the project memory
``tabledown-no-remote-telemetry`` for why remote crash reporting was rejected.

Two pieces:

  1. ``install_crash_hooks()`` — route currently-invisible failures into the
     existing ``Tabledown.log``:
       • ``sys.excepthook`` for the main thread;
       • ``threading.excepthook`` — REQUIRED, because all real work runs on the
         daemon clipboard-watcher thread, which ``sys.excepthook`` never sees, so
         an uncaught exception there dies silently today;
       • ``faulthandler`` for native faults (a segfault in PyObjC/Carbon/Pillow
         leaves no Python traceback) — dumped to a sibling ``Tabledown.crash``
         file held open for the process lifetime (async-signal-safe).

  2. ``export_diagnostics()`` + ``reveal()`` — write a SCRUBBED bundle (home
     paths, usernames, volume names, and obvious secrets removed) and reveal it
     in Finder. Touches no clipboard, so it can never race the watcher.
"""
from __future__ import annotations

import faulthandler
import re
import sys
import threading
import traceback
from pathlib import Path

from .logger import BACKUP_LOG_PATH, LOG_PATH, log
from . import __version__

# Native-fault dumps land here. The file is opened once and held in a module
# global for the whole process: faulthandler writes from a signal handler
# (async-signal-safe), so it needs a valid fd that can't be GC'd or reopened.
CRASH_PATH = LOG_PATH.with_name("Tabledown.crash")
DIAGNOSTICS_PATH = LOG_PATH.with_name("Tabledown-diagnostics.txt")
_crash_file = None  # must outlive install_crash_hooks(); do NOT make this local


# --- Crash capture --------------------------------------------------------

def install_crash_hooks() -> None:
    """Install Python + native crash capture into the local log. Best-effort."""
    _install_python_hooks()
    _install_faulthandler()


def _install_python_hooks() -> None:
    previous = sys.excepthook

    def _hook(exc_type, exc_value, exc_tb):
        _record("main", exc_type, exc_value, exc_tb)
        try:
            previous(exc_type, exc_value, exc_tb)  # keep default stderr behavior
        except Exception:  # noqa: BLE001 - a crash hook must never crash
            pass

    sys.excepthook = _hook

    # The clipboard watcher is a daemon thread; sys.excepthook is blind to it.
    if hasattr(threading, "excepthook"):
        def _thread_hook(args):
            _record(getattr(args.thread, "name", "thread"),
                    args.exc_type, args.exc_value, args.exc_traceback)
        threading.excepthook = _thread_hook


def _install_faulthandler() -> None:
    global _crash_file
    try:
        CRASH_PATH.parent.mkdir(parents=True, exist_ok=True)
        _crash_file = open(CRASH_PATH, "a", buffering=1, encoding="utf-8")
        faulthandler.enable(file=_crash_file, all_threads=True)
        import signal
        # chain=True so the OS crash reporter still runs (don't suppress the
        # .ips Apple writes). Native trampoline faults may yield only a partial
        # dump — best-effort.
        for name in ("SIGSEGV", "SIGABRT", "SIGBUS", "SIGFPE"):
            sig = getattr(signal, name, None)
            if sig is None:
                continue
            try:
                faulthandler.register(sig, file=_crash_file, all_threads=True, chain=True)
            except (RuntimeError, ValueError, OSError):
                pass
    except (OSError, ValueError, RuntimeError) as exc:
        log(f"faulthandler setup skipped: {type(exc).__name__}")


def _record(where: str, exc_type, exc_value, exc_tb) -> None:
    """Scrub + log an uncaught exception. Never raises."""
    try:
        formatted = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        log(f"UNCAUGHT [{where}]\n{scrub(formatted)}")
    except Exception:  # noqa: BLE001
        try:
            log(f"UNCAUGHT [{where}] {getattr(exc_type, '__name__', 'error')} (unformattable)")
        except Exception:  # noqa: BLE001
            pass


# --- Scrub ----------------------------------------------------------------

def scrub(text: str) -> str:
    """Remove PII so a shared log can't leak paths, usernames, volume labels, or
    secrets. Over-scrubbing is the safe direction. Home-prefix is replaced first
    so later patterns see an already-anonymized path."""
    s = text
    home = str(Path.home())
    if home and home != ".":
        s = s.replace(home, "~")
    # /Users/<name>/… and C:\Users\<name>\… (defensive: both path styles).
    s = re.sub(r"/Users/[^/\n]+/", "~/", s)
    s = re.sub(r"/Users/[^/\n]+", "~", s)
    s = re.sub(r"(?i)C:\\Users\\[^\\\n]+\\", r"%USERPROFILE%\\", s)
    # Mounted volume names are themselves private (external drive labels).
    s = re.sub(r"/Volumes/[^\"'\n]*", "/Volumes/…", s)
    # Per-user temp dirs can identify the user.
    s = re.sub(r"/private/var/folders/[^\s\"']*", "~tmp", s)
    s = re.sub(r"/var/folders/[^\s\"']*", "~tmp", s)
    user = _username()
    if user:
        s = re.sub(r"(?i)\b" + re.escape(user) + r"\b", "<user>", s)
    # Obvious secret patterns.
    s = re.sub(r"(?i)(api[_-]?key|secret|token|password|passwd|pwd|bearer)\b\s*[:=]\s*\S+",
               r"\1=***", s)
    s = re.sub(r"\b(sk|pk|ghp|gho|ghs|xox[baprs])[-_][A-Za-z0-9]{8,}", "***", s)
    return s


def _username() -> str:
    try:
        import getpass
        return getpass.getuser()
    except Exception:  # noqa: BLE001
        return ""


# --- Manual export (user-initiated bug report) ----------------------------

def export_diagnostics() -> Path | None:
    """Write a SCRUBBED diagnostics bundle for a bug report; return its path or
    None. Reads only Tabledown's own log + crash file — no clipboard access."""
    try:
        parts = [
            f"Tabledown {__version__}",
            _os_string(),
            "",
            "This file is scrubbed: home paths, usernames, volume names, and",
            "secrets are removed. It contains NO clipboard or table contents.",
            "",
            "=== recent log ===",
            scrub(_tail_text(max_chars=60000)),
        ]
        crash = _read_text(CRASH_PATH)
        if crash.strip():
            parts += ["", "=== native crash dump ===", scrub(crash)[-20000:]]
        DIAGNOSTICS_PATH.write_text("\n".join(parts) + "\n", encoding="utf-8")
        return DIAGNOSTICS_PATH
    except Exception as exc:  # noqa: BLE001
        log(f"diagnostics export failed: {type(exc).__name__}")
        return None


def reveal(path: Path) -> bool:
    """Reveal ``path`` in Finder (sandbox-safe via NSWorkspace, no subprocess)."""
    try:
        from AppKit import NSWorkspace
        from Foundation import NSURL
        NSWorkspace.sharedWorkspace().activateFileViewerSelectingURLs_(
            [NSURL.fileURLWithPath_(str(path))]
        )
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"reveal failed: {type(exc).__name__}")
        return False


def _os_string() -> str:
    try:
        import platform
        return f"macOS {platform.mac_ver()[0]} ({platform.machine()})"
    except Exception:  # noqa: BLE001
        return "macOS"


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""


def _tail_text(max_chars: int) -> str:
    """The most recent log text: rotated backup (older) then the live log."""
    text = _read_text(BACKUP_LOG_PATH) + _read_text(LOG_PATH)
    return text[-max_chars:]
