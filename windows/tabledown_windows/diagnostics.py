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
     ``Tabledown.crash`` file held open for the process lifetime. Crash records
     hold the exception TYPE + frame locations only, never the message.
  2. ``export_diagnostics()`` + ``reveal()`` — write a SCRUBBED bundle and open
     its folder in Explorer. No clipboard access, so it can't race the watcher.

``scrub()`` is kept byte-for-byte identical to the macOS port's so the two don't
drift. Importable on a non-Windows test host: ``os.startfile`` (Windows-only) is
called only inside ``reveal()`` behind try/except.
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
        # enable() installs handlers for the fatal signals (SIGSEGV/SIGFPE/
        # SIGABRT; faulthandler catches common access-violations) and dumps every
        # thread's Python traceback to our held fd, then chains to the default so
        # Windows Error Reporting still runs. (register() REFUSES these fatal
        # signals — "use enable() instead" — so there is nothing more to register.)
        faulthandler.enable(file=_crash_file, all_threads=True)
    except (OSError, ValueError, RuntimeError) as exc:
        log(f"faulthandler setup skipped: {type(exc).__name__}")


def _format_record(where: str, exc_type, exc_tb) -> str:
    """Build a payload-free crash record: exception TYPE + stack frames only.

    Deliberately omits the exception message/args (``str(exc_value)``): those can
    carry clipboard/table text or other payload — the same reason the catch sites
    log ``type(exc).__name__`` only. Frame source lines are the app's own CODE,
    not user data; ``scrub`` still anonymizes the file paths."""
    name = getattr(exc_type, "__name__", "error")
    lines = [f"UNCAUGHT [{where}] {name}"]
    for frame in traceback.extract_tb(exc_tb):
        lines.append(f"  {frame.filename}:{frame.lineno} in {frame.name}: {(frame.line or '').strip()}")
    return scrub("\n".join(lines))


def _record(where: str, exc_type, exc_value, exc_tb) -> None:
    """Log an uncaught exception (type + frames, never the message). Never raises.

    ``exc_value`` is accepted to match the hook signatures but intentionally not
    logged — its text can contain payload."""
    try:
        log(_format_record(where, exc_type, exc_tb))
    except Exception:  # noqa: BLE001 - a crash hook must never crash
        try:
            log(f"UNCAUGHT [{where}] {getattr(exc_type, '__name__', 'error')}")
        except Exception:  # noqa: BLE001
            pass


# --- Scrub ----------------------------------------------------------------

def scrub(text: str) -> str:
    """Best-effort PII redaction so a shared diagnostics file can't leak paths,
    usernames, volume/drive names, or obvious secrets. Over-redaction is the safe
    direction.

    Best-effort BY DESIGN: the real guarantee is that payload is kept out of what
    gets logged (see ``_format_record`` and the type-only catch sites), not that
    this catches every secret. Kept byte-for-byte identical to the macOS port's
    ``scrub()`` so the two don't drift."""
    s = text
    home = str(Path.home())
    if home and home not in ("", ".", "/", "\\"):
        s = s.replace(home, "~")
    # User home directories (POSIX + Windows, any drive letter).
    s = re.sub(r"/Users/[^/\n]+/", "~/", s)
    s = re.sub(r"/Users/[^/\n]+", "~", s)
    s = re.sub(r"/home/[^/\n]+/", "~/", s)
    s = re.sub(r"/home/[^/\n]+", "~", s)
    s = re.sub(r"(?i)[A-Z]:\\Users\\[^\\\n]+", "~", s)
    # Mounted volume / external-drive labels (macOS) and per-user temp dirs.
    s = re.sub(r"/Volumes/[^\"'\n]*", "/Volumes/…", s)
    s = re.sub(r"/private/var/folders/[^\s\"']*", "~tmp", s)
    s = re.sub(r"/var/folders/[^\s\"']*", "~tmp", s)
    # Standalone username token (reverse-DNS ids, dispatch-queue labels, …).
    user = _username()
    if user:
        s = re.sub(r"(?i)\b" + re.escape(user) + r"\b", "<user>", s)
    # Inline URL credentials:  scheme://user:pass@host -> scheme://***@host
    s = re.sub(r"([A-Za-z][A-Za-z0-9+.\-]*://)[^/\s:@]+:[^/\s@]+@", r"\1***@", s)
    # Secrets: bearer tokens (space form — the canonical "Authorization: Bearer
    # <tok>"), keyword=value (incl. underscored keys like aws_secret_access_key),
    # and well-known credential shapes that don't carry a keyword.
    s = re.sub(r"(?i)\bbearer\s+[A-Za-z0-9._\-]{6,}", "Bearer ***", s)
    s = re.sub(r"(?i)[\w.\-]*(?:api[_-]?key|secret|token|password|passwd|pwd|credential)[\w.\-]*\s*[:=]\s*\S+",
               "***", s)
    s = re.sub(r"\b(?:AKIA|ASIA)[0-9A-Z]{12,}", "***", s)                                     # AWS access key id
    s = re.sub(r"\bAIza[0-9A-Za-z_\-]{20,}", "***", s)                                        # Google API key
    s = re.sub(r"\beyJ[A-Za-z0-9_\-]{6,}\.[A-Za-z0-9_\-]{6,}\.[A-Za-z0-9_\-]{4,}", "***", s)  # JWT
    s = re.sub(r"\b(?:sk|pk|ghp|gho|ghs|xox[baprs])[-_][A-Za-z0-9\-]{6,}", "***", s)          # prefixed tokens
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
            "Best-effort scrubbed: home paths, usernames, volume/drive names, and",
            "obvious secrets are redacted. Crash records hold only error types and",
            "code locations — never the exception message, clipboard, or table data.",
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
