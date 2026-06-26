"""Tabledown i18n for Windows."""
from __future__ import annotations

import locale
import os

from .settings import load_settings, save_setting


LANGUAGE_KEY = "language"
SUPPORTED_LANGUAGES = ("ko", "en")
DEFAULT_LANGUAGE = "en"

TRANSLATIONS: dict[str, dict[str, str]] = {
    "ko": {
        "menu.toggle": "Tabledown 사용",
        "menu.language": "언어",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "로그인 시 자동 실행",
        "menu.diagnostics": "문제 신고용 로그 열기",
        "diagnostics.export_failed": "진단 로그를 준비하지 못했습니다. 디스크 여유 공간을 확인하고 다시 시도해 주세요.",
        "menu.help": "도움말",
        "menu.quit": "종료",
        "help.title": "Tabledown",
        "login_item.blocked_by_user": (
            "자동 실행이 작업 관리자에서 꺼져 있어 앱에서 켤 수 없습니다.\n\n"
            "작업 관리자 → ‘시작 프로그램 앱’ 탭에서 Tabledown 을 ‘사용’으로 바꿔 주세요."
        ),
        "login_item.blocked_by_policy": (
            "시스템 정책이 자동 실행을 제어하고 있어 여기서 변경할 수 없습니다.\n\n"
            "관리자에게 문의해 주세요."
        ),
        "welcome.title": "Tabledown 에 오신 것을 환영합니다!",
        "welcome.intro": "Tabledown 이 작업 표시줄 알림 영역(트레이)에서 실행 중입니다 — 표 모양 아이콘을 찾아보세요.",
        "help.message": (
            "Excel ↔ Markdown 표 변환기\n\n"
            "사용법:\n"
            "1. Excel/스프레드시트 또는 마크다운 표를 복사 (Ctrl+C)\n"
            "2. 원하는 앱에서 그대로 붙여넣기 (Ctrl+V)\n\n"
            "Excel 표를 복사하면 마크다운 에디터에서 Markdown 표로 붙고,\n"
            "Markdown 표를 복사하면 Excel에서 셀에 분리되어 붙습니다.\n\n"
            "트레이 메뉴의 ‘Tabledown 사용’ 체크 표시가 현재 상태입니다.\n"
            "아이콘에 사선이 그어져 있으면 변환이 꺼져 있는 상태입니다.\n"
            "단축키 Ctrl+Alt+T 로 변환을 켜고 끌 수 있습니다."
        ),
    },
    "en": {
        "menu.toggle": "Use Tabledown",
        "menu.language": "Language",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "Open at Login",
        "menu.diagnostics": "Open logs for bug report",
        "diagnostics.export_failed": "Couldn't prepare the diagnostics log. Check free disk space and try again.",
        "menu.help": "Help",
        "menu.quit": "Quit",
        "help.title": "Tabledown",
        "login_item.blocked_by_user": (
            "Launch at login was turned off in Task Manager, so the app can’t "
            "enable it.\n\n"
            "Open Task Manager → “Startup apps” and set Tabledown to “Enabled”."
        ),
        "login_item.blocked_by_policy": (
            "A system policy controls launch at login, so it can’t be changed "
            "here.\n\n"
            "Please contact your administrator."
        ),
        "welcome.title": "Welcome to Tabledown!",
        "welcome.intro": "Tabledown is running in your notification area (system tray) — look for the table icon.",
        "help.message": (
            "Excel ↔ Markdown table converter\n\n"
            "How to use:\n"
            "1. Copy a table from Excel/Sheets or a Markdown table (Ctrl+C)\n"
            "2. Paste in any app (Ctrl+V)\n\n"
            "Excel tables paste as Markdown in a Markdown editor,\n"
            "and Markdown tables paste into separate cells in Excel.\n\n"
            "The checkmark next to 'Use Tabledown' in the tray menu shows the current state.\n"
            "A slash through the tray icon means conversion is off.\n"
            "The global shortcut Ctrl+Alt+T toggles conversion on and off."
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
    value = load_settings().get(LANGUAGE_KEY)
    return value if value in SUPPORTED_LANGUAGES else None


def save_preferred_language(lang: str) -> None:
    """Persist user-selected language to the Windows config file."""
    if lang not in SUPPORTED_LANGUAGES:
        return
    save_setting(LANGUAGE_KEY, lang)


def resolve_language() -> str:
    """Pick language: stored preference > system locale > default."""
    return load_preferred_language() or detect_system_language()


def _locale_name() -> str:
    try:
        value = locale.getlocale()[0]
    except (TypeError, ValueError):
        return ""
    return value or ""
