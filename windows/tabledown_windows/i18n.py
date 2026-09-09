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
        "menu.fill_blanks": "빈칸을 자동 채우기",
        "menu.copy_excel_formulas": "셀 값·수식·참조를 XML로 복사",
        "menu.copy_excel_formulas_ai": "AI용 간결 복사",
        "menu.language": "언어",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "로그인 시 자동 실행",
        "menu.diagnostics": "문제 신고용 로그 열기",
        "diagnostics.export_failed": "진단 로그를 준비하지 못했습니다. 디스크 여유 공간을 확인하고 다시 시도해 주세요.",
        "formula_export.success": "셀 값·수식·참조를 XML로 복사했습니다.",
        "formula_export.success.ai": "AI용 간결 XML로 복사했습니다.",
        "formula_export.success.partial_references": "수식과 현재 결과를 복사했습니다. 일부 참조값을 가져오지 못한 부분은 복사한 내용에 표시했습니다.",
        "formula_export.success.calculation_incomplete": "수식과 현재 결과를 복사했습니다. Excel이 계산 중이어서 결과가 달라질 수 있습니다. 최신 결과가 필요하면 계산이 끝난 뒤 다시 복사하세요.",
        "formula_export.success.partial_references_and_calculation": "수식과 현재 결과를 복사했습니다. 일부 참조값이 빠져 있고 Excel도 계산 중입니다. 최신 결과가 필요하면 계산이 끝난 뒤 다시 복사하세요.",
        "formula_export.error.in_progress": "수식 XML 내보내기가 이미 진행 중입니다. 완료될 때까지 기다려 주세요.",
        "formula_export.error.clipboard_changed": "XML을 준비하는 동안 새 내용이 복사되어 보호를 위해 내보내기를 취소했습니다. 필요하면 수식 XML 복사 명령을 다시 실행해 주세요.",
        "formula_export.error.clipboard_write_failed": "수식 XML은 준비했지만 클립보드에 기록하지 못했습니다. 클립보드 내용이 완전하지 않을 수 있으니 명령을 다시 실행해 주세요.",
        "formula_export.error.output_too_large": "생성될 수식 XML이 10MB를 넘습니다. 더 작은 범위를 선택해 주세요.",
        "formula_export.error.excel_not_running": "실행 중인 Excel을 찾을 수 없습니다.",
        "formula_export.error.selection_not_range": "Excel에서 셀 범위를 선택한 뒤 다시 시도해 주세요.",
        "formula_export.error.multiple_areas": "서로 떨어진 여러 범위는 지원하지 않습니다. 하나의 연속된 범위를 선택해 주세요.",
        "formula_export.error.no_formulas": "선택한 범위에 수식이 없습니다.",
        "formula_export.error.too_large": "선택 범위가 너무 큽니다. 10,000개 이하의 셀을 선택해 주세요.",
        "formula_export.error.multiple_instances": "여러 Excel 프로세스가 실행 중입니다. 다른 Excel 창을 닫고 다시 시도해 주세요.",
        "formula_export.error.too_much_text": "선택한 셀 값과 수식 내용이 너무 큽니다. 더 작은 범위를 선택해 주세요.",
        "formula_export.error.selection_changed": "표 값과 수식을 읽는 동안 Excel 선택 영역이나 셀 내용이 변경되었습니다. 범위를 다시 선택해 재시도해 주세요.",
        "formula_export.error.com_failure": "Excel 표 값과 수식을 읽지 못했습니다. 셀 범위를 다시 선택한 뒤 시도해 주세요.",
        "formula_export.error.export_failed": "Excel 표 값과 수식을 클립보드에 복사하지 못했습니다.",
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
            "Excel 수식이 포함된 한 영역을 선택한 뒤 트레이 메뉴의\n"
            "‘셀 값·수식·참조를 XML로 복사’ 를 누르거나 Ctrl+Alt+E 를 누르면 모든 셀의 주소·값 타입·원시값·계산 상태·A1/R1C1 수식과 직접 A1 참조값이 셀 그리드 XML로 복사됩니다. 병합 계층·표시 서식은 포함하지 않으며 XML에 그 제한을 표시합니다.\n"
            "‘AI용 간결 복사’는 제목·항목명을 함께 선택하면 추정한 맥락을 표시하고, 같은 참조범위의 값은 한 번만 기록해 수식에 연결합니다.\n"
            "XML을 읽는 동안 다른 내용을 복사하면 새 클립보드를 보호하기 위해 XML 쓰기를 취소합니다.\n\n"
            "트레이 메뉴의 ‘Tabledown 사용’ 체크 표시가 현재 상태입니다.\n"
            "아이콘에 사선이 그어져 있으면 변환이 꺼져 있는 상태입니다.\n"
            "단축키 Ctrl+Alt+T 로 변환을 켜고 끌 수 있습니다."
        ),
    },
    "en": {
        "menu.toggle": "Use Tabledown",
        "menu.fill_blanks": "Auto-fill blank cells",
        "menu.copy_excel_formulas": "Copy cell values, formulas, and references as XML",
        "menu.copy_excel_formulas_ai": "Copy compact XML for AI",
        "menu.language": "Language",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "Open at Login",
        "menu.diagnostics": "Open logs for bug report",
        "diagnostics.export_failed": "Couldn't prepare the diagnostics log. Check free disk space and try again.",
        "formula_export.success": "Copied cell values, formulas, and references as XML.",
        "formula_export.success.ai": "Copied compact XML for AI.",
        "formula_export.success.partial_references": "Copied the formulas and current results. Missing referenced values are marked in the copied content.",
        "formula_export.success.calculation_incomplete": "Copied the formulas and current results. Excel is still calculating, so results may change. For the latest results, copy again after calculation finishes.",
        "formula_export.success.partial_references_and_calculation": "Copied the formulas and current results. Some referenced values are missing and Excel is still calculating. For the latest results, copy again after calculation finishes.",
        "formula_export.error.in_progress": "A formula XML export is already in progress. Wait for it to finish.",
        "formula_export.error.clipboard_changed": "New content was copied while the XML was being prepared, so the export was cancelled to protect it. Run the formula XML command again if needed.",
        "formula_export.error.clipboard_write_failed": "The formula XML was prepared but could not be written to the clipboard. The clipboard may be incomplete; run the command again.",
        "formula_export.error.output_too_large": "The generated formula XML would exceed 10 MB. Select a smaller range.",
        "formula_export.error.excel_not_running": "Couldn't find a running Excel application.",
        "formula_export.error.selection_not_range": "Select a cell range in Excel, then try again.",
        "formula_export.error.multiple_areas": "Separate ranges aren't supported. Select one contiguous range.",
        "formula_export.error.no_formulas": "The selected range doesn't contain formulas.",
        "formula_export.error.too_large": "The selected range is too large. Select 10,000 cells or fewer.",
        "formula_export.error.multiple_instances": "Multiple Excel processes are running. Close the other Excel windows and try again.",
        "formula_export.error.too_much_text": "The selected cell values and formulas contain too much text. Select a smaller range.",
        "formula_export.error.selection_changed": "The Excel selection or cell contents changed while values and formulas were being read. Select the range again, then retry.",
        "formula_export.error.com_failure": "Couldn't read the Excel table values and formulas. Select the cell range again, then retry.",
        "formula_export.error.export_failed": "Couldn't copy the Excel table values and formulas to the clipboard.",
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
            "Select one formula range in Excel, then click ‘Copy cell values, formulas, and references as XML’\n"
            "in the tray menu or press Ctrl+Alt+E to copy every cell address, value type, raw value, calculation state, A1/R1C1 formula, and direct A1 reference value as cell-grid XML. The XML explicitly says that merge hierarchy and display formatting are not included.\n"
            "For ‘Copy compact XML for AI’, include headers and item names to add inferred context; identical reference ranges are written once and linked to each formula.\n"
            "If you copy something else while Excel is being read, Tabledown cancels the XML write to protect the newer clipboard content.\n\n"
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
