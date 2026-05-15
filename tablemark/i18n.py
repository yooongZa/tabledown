"""Tabledown i18n: in-memory translation dict + macOS-aware language resolution.

Unknown keys fall back to the key string. Unknown languages fall back to English.
The user-selected language is persisted via NSUserDefaults so it survives restarts.
"""
from __future__ import annotations

from Foundation import NSLocale, NSUserDefaults

LANGUAGE_KEY = "com.tabledown.app.language"
SUPPORTED_LANGUAGES = ("ko", "en")
DEFAULT_LANGUAGE = "en"

TRANSLATIONS: dict[str, dict[str, str]] = {
    "ko": {
        "menu.toggle_on": "활성화 ✓",
        "menu.toggle_off": "비활성화",
        "menu.language": "언어",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item_on": "로그인 시 자동 실행 ✓",
        "menu.login_item_off": "로그인 시 자동 실행",
        "menu.hide_icon": "메뉴바 아이콘 숨기기",
        "menu.help": "도움말",
        "menu.quit": "종료",
        "help.title": "Tabledown",
        "help.message": (
            "Excel ↔ Markdown 표 변환기\n\n"
            "사용법:\n"
            "1. Excel/스프레드시트 또는 마크다운 표를 복사 (Cmd+C)\n"
            "2. 원하는 앱에서 그대로 붙여넣기 (Cmd+V)\n\n"
            "Excel 표를 복사하면 마크다운 에디터에서 Markdown 표로 붙고,\n"
            "Markdown 표를 복사하면 Excel에서 셀에 분리되어 붙습니다."
        ),
        "hide.alert_title": "메뉴바 아이콘을 숨겼습니다",
        "hide.alert_message": (
            "Tabledown은 계속 백그라운드에서 동작합니다.\n"
            "다시 표시하려면 앱을 종료한 뒤 다시 실행하세요."
        ),
    },
    "en": {
        "menu.toggle_on": "Enabled ✓",
        "menu.toggle_off": "Disabled",
        "menu.language": "Language",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item_on": "Open at Login ✓",
        "menu.login_item_off": "Open at Login",
        "menu.hide_icon": "Hide menu bar icon",
        "menu.help": "Help",
        "menu.quit": "Quit",
        "help.title": "Tabledown",
        "help.message": (
            "Excel ↔ Markdown table converter\n\n"
            "How to use:\n"
            "1. Copy a table from Excel/Sheets or a Markdown table (Cmd+C)\n"
            "2. Paste in any app (Cmd+V)\n\n"
            "Excel tables paste as Markdown in a Markdown editor,\n"
            "and Markdown tables paste into separate cells in Excel."
        ),
        "hide.alert_title": "Menu bar icon hidden",
        "hide.alert_message": (
            "Tabledown keeps running in the background.\n"
            "Quit and relaunch the app to show the icon again."
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
    """Return 'ko' when macOS prefers Korean, otherwise 'en'."""
    try:
        languages = NSLocale.preferredLanguages() or []
        for lang in languages:
            code = str(lang).lower()
            if code.startswith("ko"):
                return "ko"
            if code.startswith("en"):
                return "en"
    except Exception:
        pass
    return DEFAULT_LANGUAGE


def load_preferred_language() -> str | None:
    """Read user-selected language from NSUserDefaults; None when unset or invalid."""
    try:
        value = NSUserDefaults.standardUserDefaults().stringForKey_(LANGUAGE_KEY)
    except Exception:
        return None
    if not value:
        return None
    code = str(value)
    return code if code in SUPPORTED_LANGUAGES else None


def save_preferred_language(lang: str) -> None:
    """Persist user-selected language to NSUserDefaults."""
    if lang not in SUPPORTED_LANGUAGES:
        return
    try:
        NSUserDefaults.standardUserDefaults().setObject_forKey_(lang, LANGUAGE_KEY)
    except Exception:
        pass


def resolve_language() -> str:
    """Pick language: stored preference > system locale > default."""
    return load_preferred_language() or detect_system_language()
