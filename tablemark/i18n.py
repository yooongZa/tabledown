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
        "menu.toggle": "Tabledown 사용",
        "menu.copy_xml": "표를 XML로 복사",
        "xml.no_table_title": "Tabledown",
        "xml.no_table_message": "클립보드에서 표를 찾을 수 없습니다.\n먼저 Excel/스프레드시트 표나 마크다운 표를 복사한 뒤 다시 시도하세요.",
        "menu.language": "언어",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "로그인 시 자동 실행",
        "menu.help": "도움말",
        "menu.quit": "종료",
        "help.title": "Tabledown",
        "help.message": (
            "Excel ↔ Markdown 표 변환기\n\n"
            "사용법:\n"
            "1. Excel/스프레드시트 또는 마크다운 표를 복사 (Cmd+C)\n"
            "2. 원하는 앱에서 그대로 붙여넣기 (Cmd+V)\n\n"
            "Excel 표를 복사하면 마크다운 에디터에서 Markdown 표로 붙고,\n"
            "Markdown 표를 복사하면 Excel에서 셀에 분리되어 붙습니다.\n\n"
            "XML:\n"
            "• LLM 프롬프트에 넣기 좋은 XML 표를 복사하면 자동으로 Markdown·표로 변환됩니다.\n"
            "• 메뉴의 ‘표를 XML로 복사’ 를 누르면 현재 클립보드의 표가 XML로 변환됩니다.\n\n"
            "메뉴의 ‘Tabledown 사용’ 항목 왼쪽 체크 표시가 현재 상태입니다.\n"
            "체크가 켜져 있으면 변환이 동작하고, 꺼져 있으면 변환이 멈춥니다.\n"
            "메뉴바 아이콘에 사선이 그어져 있으면 변환이 꺼져 있는 상태입니다."
        ),
    },
    "en": {
        "menu.toggle": "Use Tabledown",
        "menu.copy_xml": "Copy table as XML",
        "xml.no_table_title": "Tabledown",
        "xml.no_table_message": "No table found on the clipboard.\nCopy an Excel/Sheets table or a Markdown table first, then try again.",
        "menu.language": "Language",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "Open at Login",
        "menu.help": "Help",
        "menu.quit": "Quit",
        "help.title": "Tabledown",
        "help.message": (
            "Excel ↔ Markdown table converter\n\n"
            "How to use:\n"
            "1. Copy a table from Excel/Sheets or a Markdown table (Cmd+C)\n"
            "2. Paste in any app (Cmd+V)\n\n"
            "Excel tables paste as Markdown in a Markdown editor,\n"
            "and Markdown tables paste into separate cells in Excel.\n\n"
            "XML:\n"
            "• Copy an LLM-friendly XML table and it is converted to Markdown/table automatically.\n"
            "• Click ‘Copy table as XML’ to turn the clipboard table into XML.\n\n"
            "The checkmark next to ‘Use Tabledown’ shows the current state.\n"
            "When checked, conversion is on. When unchecked, conversion pauses.\n"
            "A slash through the menu bar icon means conversion is off."
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
