"""Small file logger for Tabledown diagnostics."""
from datetime import datetime
from pathlib import Path


LOG_PATH = Path.home() / "Library" / "Logs" / "Tabledown.log"
MAX_LOG_BYTES = 1_000_000
BACKUP_LOG_PATH = LOG_PATH.with_name(f"{LOG_PATH.name}.1")


def log(message: str) -> None:
    try:
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        _rotate_if_needed()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(f"{timestamp} {message}\n")
    except OSError:
        pass


def _rotate_if_needed() -> None:
    if not LOG_PATH.exists() or LOG_PATH.stat().st_size < MAX_LOG_BYTES:
        return
    if BACKUP_LOG_PATH.exists():
        BACKUP_LOG_PATH.unlink()
    LOG_PATH.replace(BACKUP_LOG_PATH)
