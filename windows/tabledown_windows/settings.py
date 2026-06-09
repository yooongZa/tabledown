"""Persisted user settings for the Windows port (JSON in %APPDATA%).

One file holds every setting (language, first-run flag, …). Writes are
read-modify-write so saving one key never clobbers the others — the original
i18n-only writer rewrote the whole file with just the language key.
"""
from __future__ import annotations

import json
import os
from pathlib import Path


def settings_path() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / "Tabledown" / "settings.json"
    return Path.home() / "AppData" / "Roaming" / "Tabledown" / "settings.json"


def load_settings() -> dict:
    """Read the settings dict; empty on missing or unreadable file."""
    path = settings_path()
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_setting(key: str, value) -> None:
    """Persist one setting, preserving every other key in the file."""
    data = load_settings()
    data[key] = value
    path = settings_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    except OSError:
        pass
