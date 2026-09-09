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
        "menu.copy_xml": "표 구조·표시값을 XML로 복사",
        "menu.copy_xml_busy": "표 구조 XML 복사 중…",
        "menu.copy_excel_formulas": "셀 값·수식·참조를 XML로 복사",
        "menu.copy_excel_formulas_busy": "수식 분석 XML 복사 중…",
        "menu.copy_ai_formulas": "AI용 간결 복사",
        "menu.copy_ai_formulas_busy": "AI용 XML 복사 중…",
        "menu.copy_ai_formulas_tooltip": "Excel에서 제목과 항목명을 함께 선택하세요. 선택 안의 제목 후보와 수식·현재값을 포함하고 반복 참조를 묶어 복사합니다.",
        "export.error_title": "Tabledown",
        "export.error.in_progress": "XML 내보내기가 이미 진행 중입니다. 완료될 때까지 기다려 주세요.",
        "table.error_title": "Tabledown",
        "table.error.excel_not_running": "Microsoft Excel이 실행 중이 아닙니다.",
        "table.error.no_selection": "Excel 워크시트에서 XML로 복사할 셀 범위를 선택한 뒤 다시 시도하세요.",
        "table.error.multiple_areas": "서로 떨어진 여러 범위는 복사할 수 없습니다. 하나의 직사각형 범위를 선택하세요.",
        "table.error.too_many_cells": "선택 범위가 너무 큽니다. 10,000개 이하의 셀을 선택하세요.",
        "table.error.too_much_text": "셀 내용이 너무 큽니다. 더 작은 범위를 선택하세요.",
        "table.error.selection_changed": "표를 읽는 동안 Excel 선택 영역이나 셀 값·병합 구조가 변경되었습니다. 범위를 다시 선택해 재시도하세요.",
        "table.error.automation_denied": "Excel Automation(자동화) 권한이 필요합니다. 시스템 설정 > 개인정보 보호 및 보안 > 자동화에서 Tabledown의 Microsoft Excel 접근을 허용하세요.",
        "table.error.partial_merge": "병합 셀 전체가 선택 범위에 포함되도록 표를 다시 선택하세요.",
        "table.error.display_overflow": "일부 값이 ##로 표시됩니다. Excel에서 전체 값이 보이도록 열 너비나 날짜·시간 값/서식을 확인한 뒤 다시 시도하세요.",
        "table.error.clipboard_changed": "XML을 준비하는 동안 새 내용이 복사되어 보호를 위해 내보내기를 취소했습니다. 필요하면 XML 복사 명령을 다시 실행하세요.",
        "table.error.clipboard_write_failed": "XML은 준비했지만 클립보드에 기록하지 못했습니다. 클립보드 내용이 완전하지 않을 수 있으니 XML 복사 명령을 다시 실행하세요.",
        "table.error.output_too_large": "생성될 XML이 10MB를 넘습니다. 헤더를 줄이거나 더 작은 범위를 선택하세요.",
        "table.error.execution_failed": "Excel에서 선택한 표를 읽지 못했습니다. 셀 범위를 다시 선택한 뒤 재시도하세요.",
        "table.error.invalid_response": "Excel에서 선택한 표를 읽지 못했습니다. 셀 범위를 다시 선택한 뒤 재시도하세요.",
        "table.error.no_table": "헤더와 데이터 행이 포함된 Excel 표 범위를 선택한 뒤 다시 시도하세요.",
        "formula.error_title": "Tabledown",
        "formula.copy_notice_title": "수식 복사 완료",
        "formula.copy_notice.partial_references": "수식과 현재 결과를 복사했습니다. 일부 참조값을 가져오지 못한 부분은 복사한 내용에 표시했습니다.",
        "formula.copy_notice.calculation_incomplete": "수식과 현재 결과를 복사했습니다. Excel이 계산 중이어서 결과가 달라질 수 있습니다. 최신 결과가 필요하면 계산이 끝난 뒤 다시 복사하세요.",
        "formula.copy_notice.partial_references_and_calculation": "수식과 현재 결과를 복사했습니다. 일부 참조값이 빠져 있고 Excel도 계산 중입니다. 최신 결과가 필요하면 계산이 끝난 뒤 다시 복사하세요.",
        "formula.error.excel_not_running": "Microsoft Excel이 실행 중이 아닙니다.",
        "formula.error.no_selection": "Excel 워크시트에서 셀 범위를 선택한 뒤 다시 시도하세요.",
        "formula.error.multiple_areas": "서로 떨어진 여러 범위는 복사할 수 없습니다. 하나의 직사각형 범위를 선택하세요.",
        "formula.error.too_many_cells": "선택 범위가 너무 큽니다. 10,000개 이하의 셀을 선택하세요.",
        "formula.error.too_fragmented": "수식 셀이 너무 많이 흩어져 있습니다. 더 작거나 연속된 범위를 선택하세요.",
        "formula.error.too_much_text": "셀 값과 수식 내용이 너무 큽니다. 더 작은 범위를 선택하세요.",
        "formula.error.no_formulas": "선택한 범위에 수식이 없습니다.",
        "formula.error.selection_changed": "표 값과 수식을 읽는 동안 Excel 선택 영역이나 셀 값·수식·참조값이 변경되었습니다. 범위를 다시 선택해 재시도하세요.",
        "formula.error.automation_denied": "Excel Automation(자동화) 권한이 필요합니다. 시스템 설정 > 개인정보 보호 및 보안 > 자동화에서 Tabledown의 Microsoft Excel 접근을 허용하세요.",
        "formula.error.clipboard_changed": "XML을 준비하는 동안 새 내용이 복사되어 보호를 위해 내보내기를 취소했습니다. 필요하면 수식 XML 복사 명령을 다시 실행하세요.",
        "formula.error.clipboard_write_failed": "수식 XML은 준비했지만 클립보드에 기록하지 못했습니다. 클립보드 내용이 완전하지 않을 수 있으니 명령을 다시 실행하세요.",
        "formula.error.output_too_large": "생성될 수식 XML이 10MB를 넘습니다. 더 작은 범위를 선택하세요.",
        "formula.error.execution_failed": "Excel에서 표 값과 수식을 읽지 못했습니다. 셀 범위를 다시 선택한 뒤 재시도하세요.",
        "formula.error.invalid_response": "Excel에서 표 값과 수식을 읽지 못했습니다. 셀 범위를 다시 선택한 뒤 재시도하세요.",
        "menu.fill_blanks": "빈칸을 자동 채우기",
        "menu.fill_blanks_tooltip": "표를 변환할 때(마크다운·XML) 병합·빈 칸을 바로 위/좌측의 칸 값으로 자동 채웁니다. 헤더 영역만 채우고 데이터(값) 영역의 빈 칸은 그대로 둡니다.",
        "menu.settings": "설정",
        "menu.language": "언어",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "로그인 시 자동 실행",
        "menu.diagnostics": "문제 신고용 로그 열기",
        "diagnostics.export_failed": "진단 로그를 준비하지 못했습니다. 디스크 여유 공간을 확인하고 다시 시도해 주세요.",
        "menu.help": "도움말",
        "menu.quit": "종료",
        "help.open_github": "GitHub 열기",
        "welcome.title": "Tabledown 에 오신 것을 환영합니다!",
        "welcome.intro": "Tabledown 이 메뉴바에서 실행 중입니다 — 화면 오른쪽 위의 표 모양 아이콘을 찾아보세요.",
        "help.message": (
            "Excel ↔ Markdown 표 변환기\n\n"
            "사용법:\n"
            "1. Excel/스프레드시트 또는 마크다운 표를 복사 (Cmd+C)\n"
            "2. 원하는 앱에서 그대로 붙여넣기 (Cmd+V)\n\n"
            "Excel 표를 복사하면 마크다운 에디터에서 Markdown 표로 붙고,\n"
            "Markdown 표를 복사하면 Excel에서 셀에 분리되어 붙습니다.\n"
            "표가 변환되면 메뉴바 아이콘이 잠깐 체크 표시로 바뀝니다.\n\n"
            "XML:\n"
            "• Excel에서 표 영역 하나를 선택한 뒤 ‘표 구조·표시값을 XML로 복사’ 를 누르거나 ⌘⌃X 를 누릅니다.\n"
            "• Cmd+C 없이 표시값·빈칸·실제 병합 범위·출처를 LLM 친화적 XML로 복사합니다. 헤더 행까지 선택하세요. 수식 자체는 포함하지 않습니다.\n"
            "• ‘빈칸을 자동 채우기’ 를 켜면 병합 없이 비워둔 그룹 열(직급 등)의 빈칸을 바로 위 값으로 채웁니다.\n\n"
            "Excel 수식:\n"
            "• Excel에서 수식이 포함된 한 영역을 선택한 뒤 ‘셀 값·수식·참조를 XML로 복사’ 를 누르거나 ⌘⌃E 를 누릅니다.\n"
            "• Cmd+C 없이 모든 셀의 주소·값 타입·원시값·계산 상태·A1/R1C1 수식과 직접 A1 참조값을 셀 그리드 XML로 복사합니다. 병합 계층·표시 서식은 포함하지 않으며 XML에 그 제한을 표시합니다.\n"
            "• XML을 읽는 동안 메뉴에 ‘복사 중…’이 표시됩니다. 그사이 다른 내용을 복사하면 새 클립보드를 보호하기 위해 XML 쓰기를 취소합니다.\n\n"
            "AI용 간결 복사:\n"
            "• 제목과 항목명을 함께 선택한 뒤 ‘AI용 간결 복사’를 누르세요. 선택 안의 제목 후보를 추정으로 표시하고, 여러 수식이 공유하는 참조값은 한 번 기록합니다. 원본 수식·현재 결과·참조 누락 표시는 유지합니다.\n\n"
            "메뉴의 ‘Tabledown 사용’ 항목 왼쪽 체크 표시가 현재 상태입니다.\n"
            "체크가 켜져 있으면 변환이 동작하고, 꺼져 있으면 변환이 멈춥니다.\n"
            "메뉴바 아이콘에 사선이 그어져 있으면 변환이 꺼져 있는 상태입니다.\n"
            "단축키 ⌘⌃T 로 변환을 켜고 끌 수 있습니다."
        ),
    },
    "en": {
        "menu.toggle": "Use Tabledown",
        "menu.copy_ai_formulas": "Copy compact XML for AI",
        "menu.copy_ai_formulas_busy": "Copying XML for AI…",
        "menu.copy_ai_formulas_tooltip": "Include headers and item names in your Excel selection. Copies title candidates, formulas, and current values with shared references listed once.",
        "menu.copy_xml": "Copy table structure and displayed values as XML",
        "menu.copy_xml_busy": "Copying table-structure XML…",
        "menu.copy_excel_formulas": "Copy cell values, formulas, and references as XML",
        "menu.copy_excel_formulas_busy": "Copying formula-analysis XML…",
        "export.error_title": "Tabledown",
        "export.error.in_progress": "An XML export is already in progress. Wait for it to finish.",
        "table.error_title": "Tabledown",
        "table.error.excel_not_running": "Microsoft Excel isn't running.",
        "table.error.no_selection": "Select the Excel cell range you want to copy as XML, then try again.",
        "table.error.multiple_areas": "Disjoint ranges aren't supported. Select one rectangular range.",
        "table.error.too_many_cells": "The selection is too large. Select 10,000 cells or fewer.",
        "table.error.too_much_text": "The cell content is too large. Select a smaller range.",
        "table.error.selection_changed": "The Excel selection, cell values, or merge structure changed while the table was being read. Select the range again and retry.",
        "table.error.automation_denied": "Excel Automation permission is required. In System Settings > Privacy & Security > Automation, allow Tabledown to access Microsoft Excel.",
        "table.error.partial_merge": "Select the table again so every merged cell is fully inside the selection.",
        "table.error.display_overflow": "Some values appear as ## in Excel. Widen the columns or fix the date/time values or formats until the full values are visible, then try again.",
        "table.error.clipboard_changed": "New content was copied while the XML was being prepared, so the export was cancelled to protect it. Run the XML command again if needed.",
        "table.error.clipboard_write_failed": "The XML was prepared but could not be written to the clipboard. The clipboard may be incomplete; run the XML command again.",
        "table.error.output_too_large": "The generated XML would exceed 10 MB. Shorten the headers or select a smaller range.",
        "table.error.execution_failed": "Couldn't read the selected Excel table. Select the cell range again and retry.",
        "table.error.invalid_response": "Couldn't read the selected Excel table. Select the cell range again and retry.",
        "table.error.no_table": "Select an Excel table range containing a header and at least one data row, then try again.",
        "formula.error_title": "Tabledown",
        "formula.copy_notice_title": "Formulas copied",
        "formula.copy_notice.partial_references": "Copied formulas and current results. Some reference values could not be included; the affected formulas are marked in the copied content.",
        "formula.copy_notice.calculation_incomplete": "Copied formulas and current results. Excel is still calculating, so the results may change. For updated results, copy again after calculation finishes.",
        "formula.copy_notice.partial_references_and_calculation": "Copied formulas and current results. Some reference values are missing, and Excel is still calculating. For updated results, copy again after calculation finishes.",
        "formula.error.excel_not_running": "Microsoft Excel isn't running.",
        "formula.error.no_selection": "Select a cell range in an Excel worksheet, then try again.",
        "formula.error.multiple_areas": "Disjoint ranges aren't supported. Select one rectangular range.",
        "formula.error.too_many_cells": "The selection is too large. Select 10,000 cells or fewer.",
        "formula.error.too_fragmented": "The formula cells are too fragmented. Select a smaller or more contiguous range.",
        "formula.error.too_much_text": "The cell values and formulas are too large. Select a smaller range.",
        "formula.error.no_formulas": "The selected range contains no formulas.",
        "formula.error.selection_changed": "The Excel selection, cell values, formulas, or reference values changed while being read. Select the range again and retry.",
        "formula.error.automation_denied": "Excel Automation permission is required. In System Settings > Privacy & Security > Automation, allow Tabledown to access Microsoft Excel.",
        "formula.error.clipboard_changed": "New content was copied while the XML was being prepared, so the export was cancelled to protect it. Run the formula XML command again if needed.",
        "formula.error.clipboard_write_failed": "The formula XML was prepared but could not be written to the clipboard. The clipboard may be incomplete; run the command again.",
        "formula.error.output_too_large": "The generated formula XML would exceed 10 MB. Select a smaller range.",
        "formula.error.execution_failed": "Couldn't read table values and formulas from Excel. Select the cell range again and retry.",
        "formula.error.invalid_response": "Couldn't read table values and formulas from Excel. Select the cell range again and retry.",
        "menu.fill_blanks": "Auto-fill blank cells",
        "menu.fill_blanks_tooltip": "When converting a table (Markdown or XML), merged/blank cells are auto-filled from the cell directly above or to the left. Only header areas are filled; data (value) cells are left as-is.",
        "menu.settings": "Settings",
        "menu.language": "Language",
        "menu.language.ko": "한국어",
        "menu.language.en": "English",
        "menu.login_item": "Open at Login",
        "menu.diagnostics": "Open logs for bug report",
        "diagnostics.export_failed": "Couldn't prepare the diagnostics log. Check free disk space and try again.",
        "menu.help": "Help",
        "menu.quit": "Quit",
        "help.open_github": "Open GitHub",
        "welcome.title": "Welcome to Tabledown!",
        "welcome.intro": "Tabledown is running in your menu bar — look for the table icon near the top-right of your screen.",
        "help.message": (
            "Excel ↔ Markdown table converter\n\n"
            "How to use:\n"
            "1. Copy a table from Excel/Sheets or a Markdown table (Cmd+C)\n"
            "2. Paste in any app (Cmd+V)\n\n"
            "Excel tables paste as Markdown in a Markdown editor,\n"
            "and Markdown tables paste into separate cells in Excel.\n"
            "When a table is converted, the menu bar icon briefly shows a checkmark.\n\n"
            "XML:\n"
            "• Select one table range in Excel, then click ‘Copy table structure and displayed values as XML’ or press ⌘⌃X.\n"
            "• Without Cmd+C, it copies displayed values, blanks, exact merge ranges, and source metadata as LLM-friendly XML. Include the header row. Formula text is not included.\n"
            "• ‘Auto-fill blank cells’ fills blanks in left grouping columns (e.g. rank) from the value above.\n\n"
            "Excel formulas:\n"
            "• Select one formula range in Excel, then click ‘Copy cell values, formulas, and references as XML’ or press ⌘⌃E.\n"
            "• Without Cmd+C, it copies every cell address, value type, raw value, calculation state, A1/R1C1 formula, and direct A1 reference value as cell-grid XML. The XML explicitly says that merge hierarchy and display formatting are not included.\n"
            "• The menu shows ‘Copying…’ while Excel is being read. If you copy something else meanwhile, Tabledown cancels the XML write to protect the newer clipboard content.\n\n"
            "Compact XML for AI:\n"
            "• Include headers and item names, then click ‘Copy compact XML for AI’. Title candidates within the selection are marked as inferred, and reference values shared by formulas are listed once. Original formulas, current results, and missing-reference notices are preserved.\n\n"
            "The checkmark next to ‘Use Tabledown’ shows the current state.\n"
            "When checked, conversion is on. When unchecked, conversion pauses.\n"
            "A slash through the menu bar icon means conversion is off.\n"
            "The global shortcut ⌘⌃T toggles conversion on and off."
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
