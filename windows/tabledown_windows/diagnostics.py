"""Local-only diagnostics for Tabledown (Windows port).

Mirror of ``tablemark/diagnostics.py`` — same public API, same guarantee: NOTHING
is sent off the machine. It only captures failures that currently vanish from the
local log and lets the user *manually* export a scrubbed log for a bug report.
This keeps the shipped "no network connections, no telemetry ... nothing is sent
to any external server" promise intact (PRIVACY.md). See the project memory
``tabledown-no-remote-telemetry`` for why remote crash reporting was rejected.

  1. ``install_crash_hooks()`` — ``sys.excepthook`` + ``threading.excepthook``
     (the clipboard watcher is a daemon thread sys.excepthook never sees) +
     ``faulthandler`` for native access-violations, dumped to a sibling
     ``Tabledown.crash`` file held open for the process lifetime.
  2. ``export_diagnostics()`` + ``reveal()`` — write a SCRUBBED bundle and open
     its folder in Explorer. No clipboard access, so it can't race the watcher.

Importable on a non-Windows test host: ``os.startfile`` (Windows-only) is called
only inside ``reveal()`` behind try/except, and every signal is looked up with
``getattr`` so a missing ``SIGBUS`` degrades cleanly.
"""
from __future__ import annotations

import faulthandler
import os
import re
import sys
import threading
import traceback
from pathlib import Path

from .logger import BACKUP_LOG_PATH, LOG_PATH, log
from . import __version__

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
            previous(exc_type, exc_value, exc_tb)
        except Exception:  # noqa: BLE001 - a crash hook must never crash
            pass

    sys.excepthook = _hook

    # The clipboard watcher runs on a daemon thread; sys.excepthook is blind to
    # it. pystray menu callbacks run on the pump thread and are caught by
    # pystray, so they're wrapped at the call site instead.
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
        # Windows faulthandler supports a subset (no SIGBUS); getattr skips the
        # missing ones. chain=True keeps Windows Error Reporting in the loop.
        for name in ("SIGSEGV", "SIGABRT", "SIGFPE", "SIGBUS"):
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
    """Remove PII so a shared log can't leak paths, usernames, drive labels, or
    secrets. Over-scrubbing is the safe direction."""
    s = text
    home = str(Path.home())
    if home and home != ".":
        s = s.replace(home, "%USERPROFILE%")
    # C:\Users\<name>\… and /Users/<name>/… (defensive: both path styles).
    s = re.sub(r"(?i)([A-Z]:\\Users\\)[^\\\n]+(\\)", r"\1<user>\2", s)
    s = re.sub(r"(?i)([A-Z]:\\Users\\)[^\\\n]+", r"\1<user>", s)
    s = re.sub(r"/Users/[^/\n]+/", "~/", s)
    s = re.sub(r"/Users/[^/\n]+", "~", s)
    user = _username()
    if user:
        s = re.sub(r"(?i)\b" + re.escape(user) + r"\b", "<user>", s)
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
            f"Tabledown {__version__} (Windows)",
            _os_string(),
            "",
            "This file is scrubbed: home paths, usernames, drive labels, and",
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
    """Open the file's folder in Explorer (no subprocess; Windows-only API)."""
    try:
        os.startfile(str(Path(path).parent))  # type: ignore[attr-defined]  # noqa: S606
        return True
    except Exception as exc:  # noqa: BLE001 - non-Windows host or shell failure
        log(f"reveal failed: {type(exc).__name__}")
        return False


def _os_string() -> str:
    try:
        import platform
        return f"{platform.system()} {platform.version()} ({platform.machine()})"
    except Exception:  # noqa: BLE001
        return "Windows"


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""


def _tail_text(max_chars: int) -> str:
    """The most recent log text: rotated backup (older) then the live log."""
    text = _read_text(BACKUP_LOG_PATH) + _read_text(LOG_PATH)
    return text[-max_chars:]
