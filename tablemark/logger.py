"""Small file logger for Tabledown diagnostics."""
from datetime import datetime
from pathlib import Path


LOG_PATH = Path.home() / "Library" / "Logs" / "Tabledown.log"


def log(message: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(f"{timestamp} {message}\n")
