"""Tabledown i18n for Windows."""
from __future__ import annotations

import json
import locale
import os
from pathlib import Path


LANGUAGE_KEY = "language"
SUPPORTED_LANGUAGES = ("ko", "en")
DEFAULT_LANGUAGE = "en"

TRANSLATIONS: dict[str, dict[str, str]] = {
    "ko": {
        "menu.toggle_on": "활성화 ✓",
        "menu.toggle_off": "비활성화",
        "menu.language": "언어",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.help": "도움말",
        "menu.quit": "종료",
        "help.title": "Tabledown",
        "help.message": (
            "Excel ↔ Markdown 표 변환기\n\n"
            "사용법:\n"
            "1. Excel/스프레드시트 또는 마크다운 표를 복사 (Ctrl+C)\n"
            "2. 원하는 앱에서 그대로 붙여넣기 (Ctrl+V)\n\n"
            "Excel 표를 복사하면 마크다운 에디터에서 Markdown 표로 붙고,\n"
            "Markdown 표를 복사하면 Excel에서 셀에 분리되어 붙습니다."
        ),
    },
    "en": {
        "menu.toggle_on": "Enabled ✓",
        "menu.toggle_off": "Disabled",
        "menu.language": "Language",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.help": "Help",
        "menu.quit": "Quit",
        "help.title": "Tabledown",
        "help.message": (
            "Excel ↔ Markdown table converter\n\n"
            "How to use:\n"
            "1. Copy a table from Excel/Sheets or a Markdown table (Ctrl+C)\n"
            "2. Paste in any app (Ctrl+V)\n\n"
            "Excel tables paste as Markdown in a Markdown editor,\n"
            "and Markdown tables paste into separate cells in Excel."
        ),
    },
}


def t(key: str, lang: str) -> str:
    """Translate `key` for `lang`. Falls back to English, then the key itself."""
    primary = TRANSLATIONS.get(lang)
    if primary and key in primary:
        return primary[key]
    fallback = TRANSLATIONS[DEFAULT_LANGUAGE]
    return fallback.get(key, key)


def detect_system_language() -> str:
    """Return 'ko' when Windows prefers Korean, otherwise 'en'."""
    candidates = [
        _locale_name(),
        os.environ.get("LANG", ""),
        os.environ.get("LANGUAGE", ""),
    ]
    for candidate in candidates:
        code = (candidate or "").lower()
        if code.startswith("ko"):
            return "ko"
        if code.startswith("en"):
            return "en"
    return DEFAULT_LANGUAGE


def load_preferred_language() -> str | None:
    """Read user-selected language from the Windows config file."""
    path = _settings_path()
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    value = data.get(LANGUAGE_KEY)
    return value if value in SUPPORTED_LANGUAGES else None


def save_preferred_language(lang: str) -> None:
    """Persist user-selected language to the Windows config file."""
    if lang not in SUPPORTED_LANGUAGES:
        return
    path = _settings_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({LANGUAGE_KEY: lang}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except OSError:
        pass


def resolve_language() -> str:
    """Pick language: stored preference > system locale > default."""
    return load_preferred_language() or detect_system_language()


def _locale_name() -> str:
    try:
        value = locale.getlocale()[0]
    except (TypeError, ValueError):
        return ""
    return value or ""


def _settings_path() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        return Path(base) / "Tabledown" / "settings.json"
    return Path.home() / "AppData" / "Roaming" / "Tabledown" / "settings.json"
