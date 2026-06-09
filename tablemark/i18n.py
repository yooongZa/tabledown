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
        "menu.donate": "❤️ 개발자 응원하기",
        "menu.donate.small": "커피 한 모금 ☕",
        "menu.donate.medium": "커피 한 잔 ☕☕",
        "menu.donate.large": "든든한 후원 ❤️",
        "menu.restore": "구매 복원",
        "xml.locked_title": "XML 변환은 Pro 기능입니다",
        "xml.locked_message": "연 {price} 구독으로 XML 변환과 단축키(⌘⌃C)를 사용하세요.",
        "xml.subscribe_button": "Pro 구독 ({price}/년)",
        "pro.activated_title": "Pro 활성화 완료",
        "pro.activated_message": (
            "구독해 주셔서 감사합니다!\n"
            "이제 ‘표를 XML로 복사’ 메뉴와 단축키 ⌘⌃C 를 사용할 수 있습니다."
        ),
        "purchase.failed_title": "구매 실패",
        "purchase.failed_message": "구매를 완료하지 못했습니다.\n잠시 후 다시 시도해 주세요.",
        "restore.done": "구매가 복원되었습니다.",
        "restore.none": "복원할 구매 내역이 없습니다.",
        "restore.failed": "복원하지 못했습니다.\n잠시 후 다시 시도해 주세요.",
        "tip.thanks_title": "후원 감사합니다",
        "tip.thanks": "후원해 주셔서 감사합니다!",
        "menu.fill_blanks": "XML: 빈칸을 자동 채우기",
        "menu.fill_blanks_tooltip": "표를 XML로 복사할 때, 비어있는 칸을 바로 위/좌측의 칸 값으로 자동 채웁니다. 데이터(값) 영역의 빈 칸은 그대로 둡니다.",
        "xml.no_table_title": "Tabledown",
        "xml.no_table_message": "클립보드에서 표를 찾을 수 없습니다.\n먼저 Excel/스프레드시트 표나 마크다운 표를 복사한 뒤 다시 시도하세요.",
        "menu.settings": "설정",
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
            "• 메뉴의 ‘표를 XML로 복사’ 를 누르면 현재 클립보드의 표가 LLM 친화적 XML로 변환됩니다.\n"
            "• ‘XML: 빈칸을 자동 채우기’ 를 켜면 병합 없이 비워둔 그룹 열(직급 등)의 빈칸을 바로 위 값으로 채웁니다.\n\n"
            "메뉴의 ‘Tabledown 사용’ 항목 왼쪽 체크 표시가 현재 상태입니다.\n"
            "체크가 켜져 있으면 변환이 동작하고, 꺼져 있으면 변환이 멈춥니다.\n"
            "메뉴바 아이콘에 사선이 그어져 있으면 변환이 꺼져 있는 상태입니다."
        ),
    },
    "en": {
        "menu.toggle": "Use Tabledown",
        "menu.copy_xml": "Copy table as XML",
        "menu.donate": "❤️ Support the developer",
        "menu.donate.small": "A sip of coffee ☕",
        "menu.donate.medium": "A cup of coffee ☕☕",
        "menu.donate.large": "A generous tip ❤️",
        "menu.restore": "Restore Purchases",
        "xml.locked_title": "XML conversion is a Pro feature",
        "xml.locked_message": "Subscribe for {price}/year to use XML conversion and the ⌘⌃C shortcut.",
        "xml.subscribe_button": "Subscribe to Pro ({price}/yr)",
        "pro.activated_title": "Pro is now active",
        "pro.activated_message": (
            "Thanks for subscribing!\n"
            "You can now use 'Copy table as XML' and the ⌘⌃C shortcut."
        ),
        "purchase.failed_title": "Purchase failed",
        "purchase.failed_message": "The purchase could not be completed.\nPlease try again later.",
        "restore.done": "Your purchases have been restored.",
        "restore.none": "No purchases to restore.",
        "restore.failed": "Restore failed.\nPlease try again later.",
        "tip.thanks_title": "Thank you!",
        "tip.thanks": "Thanks for your support!",
        "menu.fill_blanks": "XML: Auto-fill blank cells",
        "menu.fill_blanks_tooltip": "When copying a table as XML, blank cells are auto-filled from the cell directly above or to the left. Data (value) cells are left as-is.",
        "xml.no_table_title": "Tabledown",
        "xml.no_table_message": "No table found on the clipboard.\nCopy an Excel/Sheets table or a Markdown table first, then try again.",
        "menu.settings": "Settings",
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
            "• Click ‘Copy table as XML’ to turn the clipboard table into LLM-friendly XML.\n"
            "• ‘XML: Auto-fill blank cells’ fills blanks in left grouping columns (e.g. rank) from the value above.\n\n"
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
