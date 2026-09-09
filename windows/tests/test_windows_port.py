from __future__ import annotations

import contextlib
from decimal import Decimal
import importlib
import sys
import threading
import types
import unittest
from pathlib import Path
from unittest import mock

from bs4 import BeautifulSoup


WINDOWS_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = WINDOWS_ROOT.parent
for path in (PROJECT_ROOT, WINDOWS_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from tabledown_windows import excel_formula, hotkey, single_instance, startup_task
from tabledown_windows.conversion import WINDOWS_DROP_FORMATS, converted_clipboard
from tabledown_windows.html_clipboard import CF_HTML_FORMAT_NAME, build_cf_html, extract_cf_html
from tabledown_windows.i18n import SUPPORTED_LANGUAGES, detect_system_language, t


MARKDOWN_BASIC = "| Name | Score |\n| --- | --- |\n| Alice | 95 |"
HTML_BASIC = "<table><tr><th>Name</th><th>Score</th></tr><tr><td>Alice</td><td>95</td></tr></table>"
# A table embedded in a document (paragraphs around it) — mirrors HTML_NOISE in
# scripts/run_test_matrix.py so both ports exercise the same document case.
HTML_DOCUMENT = (
    "<html><body><p>Before</p>"
    "<table><tr><th>제품</th><th>수량</th></tr><tr><td>키보드</td><td>3</td></tr></table>"
    "<p>After</p></body></html>"
)
# A bare table with a vertical merge (부장 spans 2 rows) — mirrors HTML_MD_VKEY in
# scripts/run_test_matrix.py. The "빈칸을 자동 채우기" toggle (0.5.0) forward-fills
# the merged key column in the Markdown path; off by default the blank stays.
HTML_MERGED_VKEY = (
    "<table><tr><th>직급</th><th>이름</th><th>금액</th></tr>"
    "<tr><td rowspan='2'>부장</td><td>김철수</td><td>100</td></tr>"
    "<tr><td>이영희</td><td>200</td></tr>"
    "<tr><td>차장</td><td>박민수</td><td>300</td></tr></table>"
)
MERGED_VKEY_TEXT = "직급\t이름\t금액\n부장\t김철수\t100\n이영희\t200\n차장\t박민수\t300"


def _excel_style_cf_html(interior: str) -> bytes:
    """Build CF_HTML the way Excel/Sheets do: the StartFragment/EndFragment
    markers sit *inside* the table, so the fragment is the table interior
    (``<col>``/``<tr>``/``<td>``) with the ``<table>`` tag left outside it.
    """
    pre = "<html>\r\n<body>\r\n<table border=0>"  # <table> precedes the fragment
    post = "</table>\r\n</body>\r\n</html>"
    html = pre + interior + post

    header_tmpl = (
        "Version:1.0\r\nStartHTML:0000000000\r\nEndHTML:0000000000\r\n"
        "StartFragment:0000000000\r\nEndFragment:0000000000\r\n"
    )
    base = len(header_tmpl.encode("utf-8"))
    start_html = base
    start_fragment = base + len(pre.encode("utf-8"))  # points AFTER <table>
    end_fragment = start_fragment + len(interior.encode("utf-8"))
    end_html = base + len(html.encode("utf-8"))
    header = (
        f"Version:1.0\r\nStartHTML:{start_html:010d}\r\nEndHTML:{end_html:010d}\r\n"
        f"StartFragment:{start_fragment:010d}\r\nEndFragment:{end_fragment:010d}\r\n"
    )
    return header.encode("utf-8") + html.encode("utf-8")


class WindowsPortTests(unittest.TestCase):
    def test_excel_cf_html_fragment_without_table_tag_is_wrapped(self):
        # Excel's fragment omits the <table> tag (markers are inside it). Without
        # wrapping, table detection fails and an Excel copy never converts.
        payload = _excel_style_cf_html(
            "<col width=80><tr><td>동해물과</td><td>백두산이</td></tr>"
            "<tr><td>마르고</td><td>닳도록</td></tr>"
        )

        html = extract_cf_html(payload)

        self.assertIn("<table", html.lower())
        self.assertIn("동해물과", html)

    def test_real_excel_table_converts_to_markdown(self):
        # The end-to-end regression for the reported "Excel tables don't convert"
        # bug: a realistic Excel CF_HTML payload + its TSV text must yield a
        # Markdown table in the text slot (html kept, only rendered images dropped).
        payload = _excel_style_cf_html(
            "<col width=80 span=2>"
            "<tr><td>동해물과</td><td>백두산이</td></tr>"
            "<tr><td>마르고</td><td>닳도록</td></tr>"
        )
        result = converted_clipboard(
            {"html": extract_cf_html(payload), "text": "동해물과\t백두산이\r\n마르고\t닳도록"}
        )

        self.assertIsNotNone(result)
        self.assertIn("| 동해물과 | 백두산이 |", result["text"])
        self.assertIn("| 마르고 | 닳도록 |", result["text"])
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])

    def test_cf_html_roundtrip_preserves_utf8_fragment(self):
        payload = build_cf_html("<table><tr><td>한글</td></tr></table>")

        html = extract_cf_html(payload)

        self.assertIn("<table>", html)
        self.assertIn("한글", html)

    def test_decode_html_prefers_cp1252_over_utf16_mojibake(self):
        # An even-length, invalid-UTF-8 cp1252 fragment must decode as cp1252,
        # not be mis-read as UTF-16 (which "succeeds" and merges latin byte
        # pairs into CJK glyphs, then fails table detection — a conversion macOS
        # performs because it reads ready-made HTML and never mis-decodes).
        from tabledown_windows.html_clipboard import _decode_html

        raw = "café".encode("cp1252")  # b'caf\xe9' — invalid utf-8, even length
        self.assertEqual(_decode_html(raw), "café")

    def test_decode_html_honors_utf16_bom(self):
        # A genuine UTF-16 fragment (BOM present) is still decoded correctly.
        from tabledown_windows.html_clipboard import _decode_html

        self.assertEqual(_decode_html("한글표".encode("utf-16")), "한글표")

    def test_ensure_table_wrapper_ignores_non_row_tags(self):
        # <track>/<trail>/<tr-foo> share the "<tr" prefix but are not rows; they
        # must not be wrapped into a bogus <table>. A real <tr> still is.
        from tabledown_windows.html_clipboard import _ensure_table_wrapper

        track = "<video><track kind='subtitles'></video>"
        self.assertEqual(_ensure_table_wrapper(track), track)
        self.assertIn("<table>", _ensure_table_wrapper("<tr><td>x</td></tr>"))

    def test_bare_excel_table_augments_text_and_keeps_html(self):
        # Invariant 3 (unified in 0.2.4): a bare Excel/Sheets table gains a
        # Markdown text slot but KEEPS CF_HTML (only rendered images are
        # dropped), so re-pasting into Excel/Word still yields a real table.
        result = converted_clipboard({"html": HTML_BASIC, "text": "Name\tScore\nAlice\t95"})

        self.assertEqual(result["text"], "\n| Name | Score |\n| --- | --- |\n| Alice | 95 |\n")
        self.assertIsNone(result.get("html"))
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])
        self.assertEqual(result["drop_formats"], WINDOWS_DROP_FORMATS)

    def test_html_table_with_mismatched_markdown_text_preserved(self):
        # Invariant 1: markdown text + html <table> is a real web/chat table. A
        # mismatched separator row must NOT trigger html->md (which would strip
        # html and break Excel paste); the clipboard is left untouched.
        result = converted_clipboard(
            {"html": HTML_BASIC, "text": "| # | A | B |\n| --- | --- |\n| 1 | x | y |"}
        )

        self.assertIsNone(result)

    def test_table_in_document_augments_text_and_keeps_html(self):
        # Invariant 4: a table inside a document augments the text slot with a
        # Markdown table while keeping the surrounding text, and preserves CF_HTML.
        result = converted_clipboard(
            {"html": HTML_DOCUMENT, "text": "Before\n제품 수량\n키보드 3\nAfter"}
        )

        self.assertIn("| 제품 | 수량 |", result["text"])
        self.assertIn("Before", result["text"])
        self.assertIn("After", result["text"])
        self.assertIsNone(result.get("html"))
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])

    def test_fill_blanks_off_keeps_merged_header_blank(self):
        # 0.5.0 default (toggle OFF): a merged cell stays blank in the Markdown —
        # no invented data, behavior unchanged from before the toggle existed.
        result = converted_clipboard({"html": HTML_MERGED_VKEY, "text": MERGED_VKEY_TEXT})

        self.assertIn("|   | 이영희 | 200 |", result["text"])
        self.assertNotIn("| 부장 | 이영희 | 200 |", result["text"])

    def test_fill_blanks_on_fills_merged_header(self):
        # 0.5.0 (toggle ON): the merged key column (부장, rowspan) carries down to
        # 이영희's row instead of leaving a blank — parity with the macOS path.
        result = converted_clipboard(
            {"html": HTML_MERGED_VKEY, "text": MERGED_VKEY_TEXT}, fill_blanks=True
        )

        self.assertIn("| 부장 | 이영희 | 200 |", result["text"])
        # HTML slot is still preserved (only rendered images dropped) — the fill
        # is a text-slot concern and must not change the drop set (invariant 3).
        self.assertNotIn(CF_HTML_FORMAT_NAME, result["drop_formats"])

    def test_markdown_clipboard_adds_html_table(self):
        result = converted_clipboard({"text": MARKDOWN_BASIC})

        self.assertEqual(result["text"], MARKDOWN_BASIC)
        self.assertIn("<table><tr><th>Name</th><th>Score</th></tr>", result["html"])

    def test_markdown_unicode_survives_cf_html_bytes_and_preserves_text(self):
        markdown = "\n| 상품 | 수량 |\n| --- | --- |\n| 한글 확인 😀 & <값> | 2 |\n"
        result = converted_clipboard({"text": markdown})

        self.assertIsNotNone(result)
        self.assertEqual(result["text"], markdown)
        fragment = result["html"]
        fragment_bytes = fragment.encode("utf-8")
        payload = build_cf_html(fragment)
        offsets = {}
        for line in payload.split(b"\r\n")[1:5]:
            key, value = line.split(b":", 1)
            offsets[key.decode("ascii")] = int(value)
        self.assertGreaterEqual(offsets["StartHTML"], 0)
        self.assertLessEqual(offsets["StartHTML"], offsets["StartFragment"])
        self.assertLess(offsets["StartFragment"], offsets["EndFragment"])
        self.assertLessEqual(offsets["EndFragment"], offsets["EndHTML"])
        self.assertEqual(offsets["EndHTML"], len(payload))
        document_bytes = payload[offsets["StartHTML"]:offsets["EndHTML"]]
        self.assertTrue(document_bytes.startswith(b"<html>"))
        self.assertTrue(document_bytes.endswith(b"</html>"))
        self.assertGreater(len(fragment_bytes), len(fragment))
        self.assertEqual(
            payload[offsets["StartFragment"]:offsets["EndFragment"]], fragment_bytes
        )
        self.assertEqual(extract_cf_html(payload), fragment)

        # The declared encoding must decode the actual multibyte payload, and
        # HTML parsing must recover cell text rather than interpret it as tags.
        declaration = BeautifulSoup(fragment, "html.parser").find("meta", charset=True)
        self.assertIsNotNone(declaration)
        decoded = fragment_bytes.decode(declaration["charset"])
        table = BeautifulSoup(decoded, "html.parser").find("table")
        self.assertIsNotNone(table)
        rows = [
            [cell.get_text() for cell in row.find_all(["th", "td"], recursive=False)]
            for row in table.find_all("tr")
        ]
        self.assertEqual(rows, [["상품", "수량"], ["한글 확인 😀 & <값>", "2"]])
        self.assertIn("&amp;", fragment)
        self.assertIn("&lt;값&gt;", fragment)
        self.assertIsNone(
            converted_clipboard({"text": markdown, "html": extract_cf_html(payload)})
        )

    def test_generated_clipboard_is_skipped(self):
        self.assertIsNone(converted_clipboard({"generated": True, "text": MARKDOWN_BASIC}))

    def test_i18n_fallbacks(self):
        self.assertIn(detect_system_language(), SUPPORTED_LANGUAGES)
        self.assertEqual(t("menu.help", "ko"), "도움말")
        self.assertEqual(t("menu.help", "fr"), "Help")
        self.assertEqual(t("missing.key", "ko"), "missing.key")

    def test_login_item_translations_exist(self):
        # The menu label and both "couldn't enable" hints must resolve in each
        # language — not fall back to the bare key.
        self.assertEqual(t("menu.login_item", "ko"), "로그인 시 자동 실행")
        self.assertEqual(t("menu.login_item", "en"), "Open at Login")
        for lang in SUPPORTED_LANGUAGES:
            for key in ("login_item.blocked_by_user", "login_item.blocked_by_policy"):
                self.assertNotEqual(t(key, lang), key)

    def test_fill_blanks_translations_exist(self):
        # The 0.5.0 "빈칸을 자동 채우기" label must resolve in each language, not
        # fall back to the bare key.
        self.assertEqual(t("menu.fill_blanks", "ko"), "빈칸을 자동 채우기")
        self.assertEqual(t("menu.fill_blanks", "en"), "Auto-fill blank cells")

    def test_excel_formula_export_translations_exist(self):
        self.assertEqual(
            t("menu.copy_excel_formulas", "ko"),
            "셀 값·수식·참조를 XML로 복사",
        )
        self.assertEqual(
            t("menu.copy_excel_formulas", "en"),
            "Copy cell values, formulas, and references as XML",
        )
        for lang in SUPPORTED_LANGUAGES:
            self.assertNotEqual(t("formula_export.success", lang), "formula_export.success")
            for notice in (
                "partial_references",
                "calculation_incomplete",
                "partial_references_and_calculation",
            ):
                key = f"formula_export.success.{notice}"
                self.assertNotEqual(t(key, lang), key)
            self.assertIn("Ctrl+Alt+E", t("help.message", lang))
            for code in (
                excel_formula.EXCEL_NOT_RUNNING,
                excel_formula.SELECTION_NOT_RANGE,
                excel_formula.MULTIPLE_AREAS,
                excel_formula.NO_FORMULAS,
                excel_formula.TOO_LARGE,
                excel_formula.MULTIPLE_INSTANCES,
                excel_formula.TOO_MUCH_TEXT,
                excel_formula.SELECTION_CHANGED,
                excel_formula.COM_FAILURE,
                "in_progress",
                "clipboard_changed",
                "clipboard_write_failed",
                "output_too_large",
                "export_failed",
            ):
                key = f"formula_export.error.{code}"
                self.assertNotEqual(t(key, lang), key)


class _ComCollection:
    def __init__(self, values):
        self._values = list(values)
        self.Count = len(self._values)

    def Item(self, index):
        return self._values[index - 1]


class _FormulaCell:
    def __init__(
        self,
        address,
        *,
        formula2=None,
        formula2_r1c1=None,
        formula=None,
        formula_r1c1=None,
        has_formula=True,
    ):
        self.Address = address
        self.HasFormula = has_formula
        if formula2 is not None:
            self.Formula2 = formula2
        if formula2_r1c1 is not None:
            self.Formula2R1C1 = formula2_r1c1
        if formula is not None:
            self.Formula = formula
        if formula_r1c1 is not None:
            self.FormulaR1C1 = formula_r1c1


class _FormulaArea:
    def __init__(self, cells):
        self.Cells = _ComCollection(cells)


class _FormulaRange:
    def __init__(self, *areas):
        self.Areas = _ComCollection(areas)


class _ReferenceValueRange:
    def __init__(self, address, values, *, row=1, column=1, rows=1, columns=1):
        self.Address = address
        self.Row = row
        self.Column = column
        self.Rows = types.SimpleNamespace(Count=rows)
        self.Columns = types.SimpleNamespace(Count=columns)
        self.CountLarge = rows * columns
        self.Value2 = values


class _ReferenceWorksheet:
    def __init__(self, ranges):
        self._ranges = ranges
        self.range_calls = []

    def Range(self, address):
        self.range_calls.append(address)
        return self._ranges[address]


class _ReferenceWorksheets:
    def __init__(self, sheets):
        self._sheets = sheets

    def Item(self, name):
        return self._sheets[name]


_DEFAULT_VALUES = object()


class _ExcelSelection:
    def __init__(
        self,
        formula_range,
        *,
        count=1,
        area_count=1,
        has_formula=True,
        single_cell=None,
        row_count=1,
        column_count=None,
        start_row=2,
        start_column=2,
        values=_DEFAULT_VALUES,
    ):
        if column_count is None:
            column_count = count // row_count
        if row_count * column_count != count:
            raise ValueError("mock selection dimensions must match count")
        if single_cell is None:
            self.Areas = types.SimpleNamespace(Count=area_count)
        else:
            self.Areas = _ComCollection([_FormulaArea([single_cell])])
        self.CountLarge = count
        self.Rows = types.SimpleNamespace(Count=row_count)
        self.Columns = types.SimpleNamespace(Count=column_count)
        self.Row = start_row
        self.Column = start_column
        end_row = start_row + row_count - 1
        end_column = start_column + column_count - 1
        first = excel_formula._absolute_address(start_row, start_column)
        last = excel_formula._absolute_address(end_row, end_column)
        self.Address = first if first == last else f"{first}:{last}"
        if values is _DEFAULT_VALUES:
            if count == 1:
                values = None
            else:
                values = tuple(
                    tuple(None for _ in range(column_count))
                    for _ in range(row_count)
                )
        self._value2 = values
        self.value2_reads = 0
        workbook = types.SimpleNamespace(Name="Budget.xlsx")
        self.Worksheet = types.SimpleNamespace(Name="Summary", Parent=workbook)
        self._formula_range = formula_range
        self._has_formula = has_formula
        self.special_cells_calls = []

    @property
    def HasFormula(self):
        if isinstance(self._has_formula, BaseException):
            raise self._has_formula
        return self._has_formula

    @property
    def Value2(self):
        self.value2_reads += 1
        if isinstance(self._value2, BaseException):
            raise self._value2
        return self._value2

    def SpecialCells(self, cell_type):
        self.special_cells_calls.append(cell_type)
        if isinstance(self._formula_range, BaseException):
            raise self._formula_range
        return self._formula_range


_DEFAULT_INTERSECTION = object()


class _FakeUser32:
    def __init__(self, windows):
        self.windows = {
            hwnd: {"visible": visible, "class": class_name, "pid": pid}
            for hwnd, visible, class_name, pid in windows
        }

    @staticmethod
    def _handle_value(hwnd):
        return int(hwnd.value if hasattr(hwnd, "value") else hwnd)

    def EnumWindows(self, callback, lparam):
        for hwnd in self.windows:
            if not callback(hwnd, lparam):
                return 0
        return 1

    def IsWindowVisible(self, hwnd):
        return int(self.windows[self._handle_value(hwnd)]["visible"])

    def GetClassNameW(self, hwnd, buffer, _buffer_size):
        value = self.windows[self._handle_value(hwnd)]["class"]
        buffer.value = value
        return len(value)

    def GetWindowThreadProcessId(self, hwnd, pid_pointer):
        pid_pointer._obj.value = self.windows[self._handle_value(hwnd)]["pid"]
        return 1


class ExcelFormulaAdapterTests(unittest.TestCase):
    def _read(
        self,
        selection=None,
        *,
        get_active_error=None,
        visible_pids=frozenset({4242}),
        active_pid=4242,
        max_formula_characters=None,
        max_value_characters=None,
        intersection=_DEFAULT_INTERSECTION,
        intersect_error=None,
        final_selection=_DEFAULT_INTERSECTION,
        calculation=excel_formula.XL_CALCULATION_AUTOMATIC,
        calculation_state=excel_formula.XL_CALCULATION_STATE_DONE,
    ):
        pythoncom = mock.Mock()
        client = mock.Mock()
        if get_active_error is not None:
            client.GetActiveObject.side_effect = get_active_error
        else:
            selections = (
                [selection]
                if final_selection is _DEFAULT_INTERSECTION
                else [selection, final_selection]
            )

            class Application:
                Hwnd = 9001

                def __init__(self):
                    self.Intersect = mock.Mock()
                    self.Calculation = calculation
                    self.CalculationState = calculation_state

                @property
                def Selection(self):
                    if len(selections) > 1:
                        return selections.pop(0)
                    return selections[0]

            application = Application()
            if intersect_error is not None:
                application.Intersect.side_effect = intersect_error
            elif intersection is _DEFAULT_INTERSECTION:
                application.Intersect.side_effect = (
                    lambda candidate, _original_selection: candidate
                )
            else:
                application.Intersect.return_value = intersection
            client.GetActiveObject.return_value = application

        read_kwargs = {}
        if max_formula_characters is not None:
            read_kwargs["max_formula_characters"] = max_formula_characters
        if max_value_characters is not None:
            read_kwargs["max_value_characters"] = max_value_characters
        with mock.patch.object(
            excel_formula, "_load_com_modules", return_value=(pythoncom, client)
        ), mock.patch.object(
            excel_formula,
            "_visible_excel_process_ids",
            return_value=set(visible_pids),
        ), mock.patch.object(
            excel_formula, "_window_process_id", return_value=active_pid
        ):
            result = excel_formula.read_selected_excel_formulas(**read_kwargs)
        pythoncom.CoInitialize.assert_called_once_with()
        pythoncom.CoUninitialize.assert_called_once_with()
        return result, client

    def test_visible_excel_pid_enumeration_filters_and_deduplicates(self):
        user32 = _FakeUser32(
            [
                (101, True, "XLMAIN", 7001),
                (102, True, "xlmain", 7001),  # another workbook, same process
                (103, True, "XLMAIN", 7002),
                (104, False, "XLMAIN", 7003),
                (105, True, "Notepad", 7004),
            ]
        )

        self.assertEqual(
            excel_formula._visible_excel_process_ids(user32), {7001, 7002}
        )
        self.assertEqual(excel_formula._window_process_id(101, user32), 7001)

    def test_multiple_windows_from_one_excel_process_are_allowed(self):
        user32 = _FakeUser32(
            [
                (101, True, "XLMAIN", 7001),
                (102, True, "XLMAIN", 7001),
            ]
        )

        self.assertEqual(excel_formula._visible_excel_process_ids(user32), {7001})

    def test_reads_single_scalar_formula_and_metadata(self):
        cell = _FormulaCell(
            "$B$2",
            formula2="=A2#",
            formula2_r1c1="=RC[-1]#",
            formula="=LEGACY_A1",
            formula_r1c1="=LEGACY_R1C1",
        )
        selection = _ExcelSelection(
            _FormulaRange(_FormulaArea([cell])), single_cell=cell
        )

        result, client = self._read(selection)

        self.assertTrue(result.ok)
        self.assertEqual(result.selection.workbook, "Budget.xlsx")
        self.assertEqual(result.selection.sheet, "Summary")
        self.assertEqual(result.selection.address, "$B$2")
        self.assertEqual(result.selection.row_count, 1)
        self.assertEqual(result.selection.column_count, 1)
        self.assertEqual(result.selection.cells[0].address, "$B$2")
        self.assertEqual(result.selection.cells[0].formula_a1, "=A2#")
        self.assertEqual(result.selection.cells[0].formula_r1c1, "=RC[-1]#")
        self.assertEqual(result.selection.calculation_mode, "automatic")
        self.assertEqual(result.selection.calculation_state, "done")
        self.assertEqual(selection.special_cells_calls, [])
        client.GetActiveObject.return_value.Intersect.assert_not_called()
        client.GetActiveObject.assert_called_once_with("Excel.Application")

    def test_enriches_same_and_cross_sheet_static_reference_values(self):
        selection = excel_formula.ExcelFormulaSelection(
            "Budget.xlsx",
            "기본_표",
            "$E$6",
            1,
            1,
            (
                excel_formula.ExcelSelectionCell(
                    "$E$6",
                    "120",
                    "=C6*'단가 표'!D6",
                    "=RC[-2]*'단가 표'!R6C4",
                ),
            ),
        )
        workbook = types.SimpleNamespace(
            Worksheets=_ReferenceWorksheets(
                {
                    "기본_표": _ReferenceWorksheet(
                        {
                            "$C$6": _ReferenceValueRange(
                                "$C$6", 10, row=6, column=3
                            )
                        }
                    ),
                    "단가 표": _ReferenceWorksheet(
                        {
                            "$D$6": _ReferenceValueRange(
                                "$D$6", 12, row=6, column=4
                            )
                        }
                    ),
                }
            )
        )

        enriched = excel_formula._enrich_formula_references(
            workbook,
            selection,
            max_value_characters=100,
        )

        formula_cell = enriched.cells[0]
        self.assertTrue(formula_cell.references_complete)
        self.assertEqual(
            [
                (reference.sheet, reference.address, reference.cells[0].value)
                for reference in formula_cell.references
            ],
            [
                ("기본_표", "$C$6", "10"),
                ("단가 표", "$D$6", "12"),
            ],
        )
        self.assertEqual(
            [
                reference.cells[0].value_kind
                for reference in formula_cell.references
            ],
            ["number", "number"],
        )

    def test_native_numbers_expand_exponents_without_changing_text(self):
        cases = (
            (315000.0, "315000"),
            (1.2e-7, "0.00000012"),
            (Decimal("3.15E+5"), "315000"),
            ("3.15E+5", "3.15E+5"),
            (True, "True"),
        )

        for value, expected in cases:
            with self.subTest(value=value):
                self.assertEqual(excel_formula._value_to_text(value), expected)

    def test_calculation_metadata_is_normalized_without_triggering_calculation(self):
        excel = types.SimpleNamespace(
            Calculation=excel_formula.XL_CALCULATION_MANUAL,
            CalculationState=excel_formula.XL_CALCULATION_STATE_PENDING,
        )

        self.assertEqual(excel_formula._calculation_mode(excel), "manual")
        self.assertEqual(excel_formula._calculation_state(excel), "pending")
        self.assertEqual(
            excel_formula._calculation_mode(types.SimpleNamespace()),
            "unknown",
        )
        self.assertEqual(
            excel_formula._calculation_state(types.SimpleNamespace()),
            "unknown",
        )

    def test_shared_reference_is_read_once_and_attached_to_every_owner(self):
        cells = tuple(
            excel_formula.ExcelSelectionCell(
                f"$B${row}",
                str(row),
                "=$A$1",
            )
            for row in range(1, 10_001)
        )
        selection = excel_formula.ExcelFormulaSelection(
            "Budget.xlsx",
            "Sheet1",
            "$B$1:$B$10000",
            10_000,
            1,
            cells,
        )
        worksheet = _ReferenceWorksheet(
            {
                "$A$1": _ReferenceValueRange(
                    "$A$1",
                    7,
                    row=1,
                    column=1,
                )
            }
        )
        workbook = types.SimpleNamespace(
            Worksheets=_ReferenceWorksheets({"Sheet1": worksheet})
        )

        enriched = excel_formula._enrich_formula_references(
            workbook,
            selection,
            max_value_characters=100_000,
        )

        self.assertEqual(worksheet.range_calls, ["$A$1"])
        self.assertEqual(len(enriched.cells), 10_000)
        self.assertTrue(all(cell.references_complete for cell in enriched.cells))
        self.assertTrue(
            all(
                cell.references[0].cells[0].value == "7"
                for cell in enriched.cells
            )
        )

    def test_reference_order_is_preserved_for_each_formula_owner(self):
        selection = excel_formula.ExcelFormulaSelection(
            "Budget.xlsx",
            "Sheet1",
            "$B$1:$B$2",
            2,
            1,
            (
                excel_formula.ExcelSelectionCell("$B$1", "4", "=$A$1+$C$1"),
                excel_formula.ExcelSelectionCell("$B$2", "4", "=$C$1+$A$1"),
            ),
        )
        worksheet = _ReferenceWorksheet(
            {
                "$A$1": _ReferenceValueRange("$A$1", 1, row=1, column=1),
                "$C$1": _ReferenceValueRange("$C$1", 3, row=1, column=3),
            }
        )
        workbook = types.SimpleNamespace(
            Worksheets=_ReferenceWorksheets({"Sheet1": worksheet})
        )

        enriched = excel_formula._enrich_formula_references(
            workbook,
            selection,
            max_value_characters=100,
        )

        self.assertEqual(worksheet.range_calls, ["$A$1", "$C$1"])
        self.assertEqual(
            [reference.address for reference in enriched.cells[0].references],
            ["$A$1", "$C$1"],
        )
        self.assertEqual(
            [reference.address for reference in enriched.cells[1].references],
            ["$C$1", "$A$1"],
        )

    def test_shared_target_counts_once_at_unique_range_limit(self):
        selection = excel_formula.ExcelFormulaSelection(
            "Budget.xlsx",
            "Sheet1",
            "$B$1:$B$3",
            3,
            1,
            (
                excel_formula.ExcelSelectionCell("$B$1", "1", "=$A$1"),
                excel_formula.ExcelSelectionCell("$B$2", "1", "=$A$1"),
                excel_formula.ExcelSelectionCell("$B$3", "2", "=$C$1"),
            ),
        )
        worksheet = _ReferenceWorksheet(
            {
                "$A$1": _ReferenceValueRange("$A$1", 1, row=1, column=1),
                "$C$1": _ReferenceValueRange("$C$1", 2, row=1, column=3),
            }
        )
        workbook = types.SimpleNamespace(
            Worksheets=_ReferenceWorksheets({"Sheet1": worksheet})
        )

        with mock.patch.object(excel_formula, "MAX_REFERENCE_RANGES", 1):
            enriched = excel_formula._enrich_formula_references(
                workbook,
                selection,
                max_value_characters=100,
            )

        self.assertEqual(worksheet.range_calls, ["$A$1"])
        self.assertTrue(enriched.cells[0].references_complete)
        self.assertTrue(enriched.cells[1].references_complete)
        self.assertFalse(enriched.cells[2].references_complete)
        self.assertEqual(
            enriched.cells[2].reference_issues,
            ("range_count_limit",),
        )

    def test_reference_value_text_limit_is_not_reported_as_read_failure(self):
        selection = excel_formula.ExcelFormulaSelection(
            "Budget.xlsx",
            "Sheet1",
            "$B$1",
            1,
            1,
            (
                excel_formula.ExcelSelectionCell("$B$1", "0", "=$A$1"),
            ),
        )
        worksheet = _ReferenceWorksheet(
            {
                "$A$1": _ReferenceValueRange(
                    "$A$1",
                    "12345",
                    row=1,
                    column=1,
                )
            }
        )
        workbook = types.SimpleNamespace(
            Worksheets=_ReferenceWorksheets({"Sheet1": worksheet})
        )

        enriched = excel_formula._enrich_formula_references(
            workbook,
            selection,
            max_value_characters=4,
        )

        formula_cell = enriched.cells[0]
        self.assertEqual(formula_cell.references, ())
        self.assertFalse(formula_cell.references_complete)
        self.assertEqual(formula_cell.reference_issues, ("value_size_limit",))

    def test_dynamic_or_unavailable_reference_keeps_formula_and_marks_partial(self):
        selection = excel_formula.ExcelFormulaSelection(
            "Budget.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (
                excel_formula.ExcelSelectionCell(
                    "$A$1", "10", '=INDIRECT("Sheet2!A1")+B1'
                ),
            ),
        )
        workbook = types.SimpleNamespace(
            Worksheets=_ReferenceWorksheets({"Sheet1": _ReferenceWorksheet({})})
        )

        enriched = excel_formula._enrich_formula_references(
            workbook,
            selection,
            max_value_characters=100,
        )

        self.assertEqual(enriched.cells[0].formula_a1, '=INDIRECT("Sheet2!A1")+B1')
        self.assertEqual(enriched.cells[0].references, ())
        self.assertFalse(enriched.cells[0].references_complete)
        self.assertEqual(
            enriched.cells[0].reference_issues,
            ("dynamic_reference", "read_failed"),
        )

    def test_reference_coordinate_read_failure_keeps_original_formula_and_result(self):
        class FailingCoordinateRange(_ReferenceValueRange):
            def __getattribute__(self, name):
                if name == object.__getattribute__(self, "failing_property"):
                    raise RuntimeError("simulated COM property failure")
                return super().__getattribute__(name)

        for property_name in ("Row", "Column"):
            with self.subTest(property=property_name):
                reference = object.__new__(FailingCoordinateRange)
                reference.failing_property = property_name
                _ReferenceValueRange.__init__(reference, "$A$1", 42)
                cell = _FormulaCell(
                    "$B$2", formula2="=A1", formula2_r1c1="=R[-1]C[-1]"
                )
                selection = _ExcelSelection(
                    _FormulaRange(_FormulaArea([cell])),
                    single_cell=cell,
                    values=42,
                )
                selection.Worksheet.Parent.Worksheets = _ReferenceWorksheets(
                    {"Summary": _ReferenceWorksheet({"$A$1": reference})}
                )

                result, _ = self._read(selection)

                self.assertTrue(result.ok)
                exported_cell = result.selection.cells[0]
                self.assertEqual(exported_cell.value, "42")
                self.assertEqual(exported_cell.formula_a1, "=A1")
                self.assertEqual(exported_cell.formula_r1c1, "=R[-1]C[-1]")
                self.assertEqual(exported_cell.references, ())
                self.assertFalse(exported_cell.references_complete)
                self.assertEqual(exported_cell.reference_issues, ("read_failed",))

    def test_reads_complete_table_row_major_with_values_blanks_and_formulas(self):
        formula = _FormulaCell(
            "$D$2",
            formula2="=B2+C2",
            formula2_r1c1="=RC[-2]+RC[-1]",
        )
        selection = _ExcelSelection(
            _FormulaRange(_FormulaArea([formula])),
            count=6,
            row_count=2,
            column_count=3,
            values=((10, 20.5, 30.5), (None, "", True)),
        )

        result, _client = self._read(selection)

        self.assertTrue(result.ok)
        self.assertEqual(result.selection.row_count, 2)
        self.assertEqual(result.selection.column_count, 3)
        self.assertEqual(
            [cell.address for cell in result.selection.cells],
            ["$B$2", "$C$2", "$D$2", "$B$3", "$C$3", "$D$3"],
        )
        self.assertEqual(
            [cell.value for cell in result.selection.cells],
            ["10", "20.5", "30.5", None, "", "True"],
        )
        self.assertEqual(
            [cell.value_kind for cell in result.selection.cells],
            ["number", "number", "number", "blank", "text", "boolean"],
        )
        self.assertEqual(result.selection.cells[2].formula_a1, "=B2+C2")
        self.assertEqual(
            result.selection.cells[2].formula_r1c1, "=RC[-2]+RC[-1]"
        )
        self.assertIsNone(result.selection.cells[0].formula_a1)
        self.assertEqual(selection.value2_reads, 1)

    def test_excel_error_hresult_is_exported_as_readable_value(self):
        cell = _FormulaCell("$B$2", formula2="=1/0", formula2_r1c1="=1/0")
        div_zero_hresult = (0x800A0000 | 2007) - (1 << 32)
        selection = _ExcelSelection(
            _FormulaRange(_FormulaArea([cell])),
            single_cell=cell,
            values=div_zero_hresult,
        )

        result, _client = self._read(selection)

        self.assertTrue(result.ok)
        self.assertEqual(result.selection.cells[0].value, "#DIV/0!")

    def test_ordinary_negative_integer_is_not_treated_as_excel_error(self):
        cell = _FormulaCell("$B$2", formula2="=-2007", formula2_r1c1="=-2007")
        selection = _ExcelSelection(
            _FormulaRange(_FormulaArea([cell])),
            single_cell=cell,
            values=-2007,
        )

        result, _client = self._read(selection)

        self.assertTrue(result.ok)
        self.assertEqual(result.selection.cells[0].value, "-2007")

    def test_malformed_bulk_value_shape_fails_closed(self):
        formula = _FormulaCell("$B$2", formula2="=1")
        selection = _ExcelSelection(
            _FormulaRange(_FormulaArea([formula])),
            count=4,
            row_count=2,
            column_count=2,
            values=((1, 2, 3), (4, 5, 6)),
        )

        result, _client = self._read(selection)

        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)
        self.assertEqual(selection.value2_reads, 1)

    def test_value_character_limit_is_content_free_and_stops_export(self):
        formula = _FormulaCell("$B$2", formula2="=1")
        selection = _ExcelSelection(
            _FormulaRange(_FormulaArea([formula])),
            count=2,
            values=(("abcd", "efgh"),),
        )

        result, _client = self._read(selection, max_value_characters=7)

        self.assertEqual(result.code, excel_formula.TOO_MUCH_TEXT)
        self.assertIsNone(result.selection)

    def test_no_formula_stops_before_bulk_values_are_read(self):
        selection = _ExcelSelection(RuntimeError("no cells"), count=2)

        result, _client = self._read(selection)

        self.assertEqual(result.code, excel_formula.NO_FORMULAS)
        self.assertEqual(selection.value2_reads, 0)

    def test_single_cell_hasformula_never_reads_specialcells_candidates(self):
        class OutsideSelectionCell:
            Address = "$Z$99"
            formula2_reads = 0

            @property
            def Formula2(self):
                self.formula2_reads += 1
                return "=OUTSIDE_SELECTION"

        outside = OutsideSelectionCell()
        selected = _FormulaCell("$B$2", formula2="=A2", formula2_r1c1="=RC[-1]")
        candidates = _FormulaRange(_FormulaArea([outside]))
        selection = _ExcelSelection(
            candidates, has_formula=True, single_cell=selected
        )

        result, client = self._read(selection)

        self.assertTrue(result.ok)
        self.assertEqual([cell.address for cell in result.selection.cells], ["$B$2"])
        self.assertEqual(selection.special_cells_calls, [])
        client.GetActiveObject.return_value.Intersect.assert_not_called()
        self.assertEqual(outside.formula2_reads, 0)

    def test_single_nonformula_cell_stops_without_specialcells_or_intersect(self):
        class OutsideSelectionCell:
            Address = "$Z$99"
            formula2_reads = 0

            @property
            def Formula2(self):
                self.formula2_reads += 1
                return "=OUTSIDE_SELECTION"

        outside = OutsideSelectionCell()
        candidates = _FormulaRange(_FormulaArea([outside]))
        selection = _ExcelSelection(candidates, has_formula=False)

        result, client = self._read(selection)

        self.assertEqual(result.code, excel_formula.NO_FORMULAS)
        self.assertEqual(selection.special_cells_calls, [])
        client.GetActiveObject.return_value.Intersect.assert_not_called()
        self.assertEqual(outside.formula2_reads, 0)

    def test_single_cell_hasformula_property_failure_is_com_error(self):
        selection = _ExcelSelection(
            None, has_formula=RuntimeError("content-free failure")
        )

        result, client = self._read(selection)

        self.assertEqual(result.code, excel_formula.COM_FAILURE)
        self.assertIsNone(result.selection)
        self.assertEqual(selection.special_cells_calls, [])
        client.GetActiveObject.return_value.Intersect.assert_not_called()

    def test_formula_to_constant_race_fails_after_a1_read(self):
        class RacingCell:
            Address = "$B$2"
            Formula2 = "=CONSTANT_THAT_LOOKS_LIKE_A_FORMULA"

            def __init__(self):
                self.states = iter((True, False))

            @property
            def HasFormula(self):
                return next(self.states)

        cell = RacingCell()
        selection = _ExcelSelection(None, single_cell=cell)

        result, _client = self._read(selection)

        # Prefix inspection would accept this value; the post-read COM state
        # check rejects it because the cell is no longer a formula.
        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)

    def test_formula_to_constant_race_fails_after_r1c1_read(self):
        class RacingCell:
            Address = "$B$2"
            Formula2 = "=A2"
            Formula2R1C1 = "=CONSTANT_R1C1"

            def __init__(self):
                self.states = iter((True, True, False))

            @property
            def HasFormula(self):
                return next(self.states)

        cell = RacingCell()
        selection = _ExcelSelection(None, single_cell=cell)

        result, _client = self._read(selection)

        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)

    def test_reads_formula_cells_from_every_specialcells_area(self):
        first = _FormulaCell("$B$2", formula2="=A2", formula2_r1c1="=RC[-1]")
        second = _FormulaCell("$D$4", formula2="=SUM(A4:C4)", formula2_r1c1="=SUM(RC[-3]:RC[-1])")
        formula_range = _FormulaRange(_FormulaArea([first]), _FormulaArea([second]))

        result, _client = self._read(
            _ExcelSelection(formula_range, count=9, row_count=3, column_count=3)
        )

        self.assertEqual(
            [
                cell.address
                for cell in result.selection.cells
                if cell.formula_a1 is not None
            ],
            ["$B$2", "$D$4"],
        )

    def test_specialcells_candidates_are_intersected_with_original_selection(self):
        class OutsideSelectionCell:
            Address = "$Z$99"
            formula2_reads = 0

            @property
            def Formula2(self):
                self.formula2_reads += 1
                return "=OUTSIDE_SELECTION"

        outside = OutsideSelectionCell()
        inside = _FormulaCell("$B$2", formula2="=A2", formula2_r1c1="=RC[-1]")
        candidates = _FormulaRange(_FormulaArea([outside, inside]))
        intersection = _FormulaRange(_FormulaArea([inside]))
        selection = _ExcelSelection(candidates, count=2)

        result, client = self._read(selection, intersection=intersection)

        self.assertTrue(result.ok)
        self.assertEqual(
            [
                cell.address
                for cell in result.selection.cells
                if cell.formula_a1 is not None
            ],
            ["$B$2"],
        )
        self.assertEqual(outside.formula2_reads, 0)
        self.assertEqual(
            client.GetActiveObject.return_value.Intersect.call_args_list,
            [mock.call(candidates, selection), mock.call(candidates, selection)],
        )

    def test_formula_topology_change_during_read_fails_closed(self):
        first = _FormulaCell("$B$2", formula2="=A2")
        second = _FormulaCell("$C$2", formula2="=A2+B2")
        initial = _FormulaRange(_FormulaArea([first, second]))
        final = _FormulaRange(_FormulaArea([first]))
        selection = _ExcelSelection(initial, count=2)
        selection.SpecialCells = mock.Mock(side_effect=[initial, final])

        result, _client = self._read(selection)

        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)
        self.assertEqual(selection.SpecialCells.call_count, 2)

    def test_active_selection_identity_change_during_read_fails_closed(self):
        cell = _FormulaCell("$B$2", formula2="=A2")
        initial = _ExcelSelection(None, single_cell=cell)
        final = _ExcelSelection(None, single_cell=cell)
        final.Address = "$C$3"

        result, _client = self._read(initial, final_selection=final)

        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)

    def test_no_specialcells_intersection_is_no_formulas(self):
        candidate = _FormulaRange(
            _FormulaArea([_FormulaCell("$Z$99", formula2="=Z1")])
        )

        result, _client = self._read(
            _ExcelSelection(candidate, count=2), intersection=None
        )

        self.assertEqual(result.code, excel_formula.NO_FORMULAS)
        self.assertIsNone(result.selection)

    def test_intersect_failure_is_content_free_com_error(self):
        candidate = _FormulaRange(
            _FormulaArea([_FormulaCell("$B$2", formula2="=A2")])
        )

        result, _client = self._read(
            _ExcelSelection(candidate, count=2),
            # Multi-cell path is the one that must fail closed on Intersect.
            # A single cell never calls Intersect at all.
            intersect_error=RuntimeError("formula content must not escape"),
        )

        self.assertEqual(result.code, excel_formula.COM_FAILURE)
        self.assertIsNone(result.selection)

    def test_falls_back_to_legacy_formula_properties(self):
        cell = _FormulaCell(
            "$B$2", formula="=A3+B3", formula_r1c1="=RC[-2]+RC[-1]"
        )

        result, _client = self._read(
            _ExcelSelection(_FormulaRange(_FormulaArea([cell])), count=2)
        )

        self.assertEqual(result.selection.cells[0].formula_a1, "=A3+B3")
        self.assertEqual(result.selection.cells[0].formula_r1c1, "=RC[-2]+RC[-1]")

    def test_empty_required_formula_is_com_error(self):
        cell = _FormulaCell("$B$2", formula2="", formula="")
        selection = _ExcelSelection(None, single_cell=cell)

        result, _client = self._read(selection)

        self.assertEqual(result.code, excel_formula.COM_FAILURE)
        self.assertIsNone(result.selection)

    def test_empty_optional_r1c1_is_omitted(self):
        cell = _FormulaCell(
            "$B$2", formula2="=A2", formula2_r1c1="", formula_r1c1=""
        )
        selection = _ExcelSelection(None, single_cell=cell)

        result, _client = self._read(selection)

        self.assertTrue(result.ok)
        self.assertIsNone(result.selection.cells[0].formula_r1c1)

    def test_multiple_visible_excel_processes_stop_before_getactiveobject(self):
        result, client = self._read(visible_pids={1001, 1002})

        self.assertEqual(result.code, excel_formula.MULTIPLE_INSTANCES)
        self.assertIsNone(result.selection)
        client.GetActiveObject.assert_not_called()

    def test_hidden_only_excel_is_not_read_through_getactiveobject(self):
        result, client = self._read(visible_pids=set())

        self.assertEqual(result.code, excel_formula.EXCEL_NOT_RUNNING)
        client.GetActiveObject.assert_not_called()

    def test_getactiveobject_window_pid_must_match_visible_excel(self):
        result, _client = self._read(visible_pids={1001}, active_pid=2002)

        self.assertEqual(result.code, excel_formula.MULTIPLE_INSTANCES)
        self.assertIsNone(result.selection)

    def test_formula_character_limit_counts_a1_and_r1c1_and_stops_early(self):
        class NeverReadCell:
            Address = "$D$2"
            formula2_reads = 0

            @property
            def Formula2(self):
                self.formula2_reads += 1
                return "=SHOULD_NOT_BE_READ"

        never_read = NeverReadCell()
        first = _FormulaCell("$B$2", formula2="=12", formula2_r1c1="=R1")
        second = _FormulaCell("$C$2", formula2="=34", formula2_r1c1="=R")
        formula_range = _FormulaRange(_FormulaArea([first, second, never_read]))

        result, _client = self._read(
            _ExcelSelection(formula_range, count=3), max_formula_characters=10
        )

        # First cell contributes 6 characters; the second pushes the combined
        # A1 + R1C1 total to 11 and must stop before reading the third cell.
        self.assertEqual(result.code, excel_formula.TOO_MUCH_TEXT)
        self.assertIsNone(result.selection)
        self.assertEqual(never_read.formula2_reads, 0)

    def test_oversized_a1_stops_before_reading_r1c1(self):
        class OversizedA1Cell:
            Address = "$B$2"
            Formula2 = "=1234"
            HasFormula = True
            r1c1_reads = 0

            @property
            def Formula2R1C1(self):
                self.r1c1_reads += 1
                return "=R1C1"

        cell = OversizedA1Cell()
        formula_range = _FormulaRange(_FormulaArea([cell]))

        result, _client = self._read(
            _ExcelSelection(formula_range, single_cell=cell),
            max_formula_characters=4,
        )

        self.assertEqual(result.code, excel_formula.TOO_MUCH_TEXT)
        self.assertEqual(cell.r1c1_reads, 0)

    def test_distinct_error_codes_have_no_selection_payload(self):
        no_range = types.SimpleNamespace()
        multiple = _ExcelSelection(None, area_count=2)
        no_formulas = _ExcelSelection(RuntimeError("no cells"), count=2)
        too_large = _ExcelSelection(None, count=excel_formula.MAX_SELECTION_CELLS + 1)
        broken_metadata = _ExcelSelection(_FormulaRange(_FormulaArea([])))
        del broken_metadata.Worksheet

        cases = (
            (
                excel_formula.EXCEL_NOT_RUNNING,
                lambda: self._read(get_active_error=RuntimeError("not running"))[0],
            ),
            (excel_formula.SELECTION_NOT_RANGE, lambda: self._read(no_range)[0]),
            (excel_formula.MULTIPLE_AREAS, lambda: self._read(multiple)[0]),
            (excel_formula.NO_FORMULAS, lambda: self._read(no_formulas)[0]),
            (excel_formula.TOO_LARGE, lambda: self._read(too_large)[0]),
            (excel_formula.COM_FAILURE, lambda: self._read(broken_metadata)[0]),
        )

        for expected, invoke in cases:
            with self.subTest(expected=expected):
                result = invoke()
                self.assertEqual(result.code, expected)
                self.assertFalse(result.ok)
                self.assertIsNone(result.selection)


class StableExcelFormulaReaderTests(unittest.TestCase):
    @staticmethod
    def _selection(value: str, formula: str, *, reference_value: str | None = None):
        references = ()
        if reference_value is not None:
            references = (
                excel_formula.ExcelFormulaReference(
                    "Sheet2",
                    "$B$2",
                    (excel_formula.ExcelReferenceCell("$B$2", reference_value),),
                ),
            )
        return excel_formula.ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (
                excel_formula.ExcelSelectionCell(
                    "$A$1",
                    value,
                    formula,
                    formula,
                    references=references,
                ),
            ),
        )

    @staticmethod
    def _success(selection):
        return excel_formula.ExcelFormulaResult(excel_formula.SUCCESS, selection)

    def test_two_equal_reads_return_a_stable_snapshot(self):
        stable = self._selection("2", "=1+1")
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[self._success(stable), self._success(stable)],
        ) as reader:
            result = excel_formula.read_stable_selected_excel_formulas()

        self.assertTrue(result.ok)
        self.assertEqual(result.selection, stable)
        self.assertEqual(reader.call_count, 2)

    def test_value_only_recalculation_retries_and_accepts_new_stable_value(self):
        old = self._selection("2", "=1+1")
        new = self._selection("3", "=1+1")
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[self._success(old), self._success(new), self._success(new)],
        ) as reader:
            result = excel_formula.read_stable_selected_excel_formulas()

        self.assertTrue(result.ok)
        self.assertEqual(result.selection, new)
        self.assertEqual(reader.call_count, 3)

    def test_formula_only_edit_retries_and_accepts_new_stable_formula(self):
        old = self._selection("2", "=1+1")
        new = self._selection("2", "=2*1")
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[self._success(old), self._success(new), self._success(new)],
        ) as reader:
            result = excel_formula.read_stable_selected_excel_formulas()

        self.assertTrue(result.ok)
        self.assertEqual(result.selection, new)
        self.assertEqual(reader.call_count, 3)

    def test_reference_value_change_retries_and_accepts_new_stable_snapshot(self):
        old = self._selection("2", "=Sheet2!B2", reference_value="1")
        new = self._selection("2", "=Sheet2!B2", reference_value="2")
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[self._success(old), self._success(new), self._success(new)],
        ) as reader:
            result = excel_formula.read_stable_selected_excel_formulas()

        self.assertTrue(result.ok)
        self.assertEqual(result.selection, new)
        self.assertEqual(reader.call_count, 3)

    def test_three_different_snapshots_fail_after_bounded_retry(self):
        snapshots = [
            self._selection("2", "=1+1"),
            self._selection("3", "=1+2"),
            self._selection("4", "=2+2"),
        ]
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[self._success(value) for value in snapshots],
        ) as reader:
            result = excel_formula.read_stable_selected_excel_formulas()

        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)
        self.assertEqual(reader.call_count, 3)

    def test_transient_selection_change_retries_but_fatal_error_does_not(self):
        stable = self._selection("2", "=1+1")
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[
                excel_formula.ExcelFormulaResult(excel_formula.SELECTION_CHANGED),
                self._success(stable),
                self._success(stable),
            ],
        ) as retrying_reader:
            result = excel_formula.read_stable_selected_excel_formulas()
        self.assertTrue(result.ok)
        self.assertEqual(retrying_reader.call_count, 3)

        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            return_value=excel_formula.ExcelFormulaResult(excel_formula.COM_FAILURE),
        ) as fatal_reader:
            result = excel_formula.read_stable_selected_excel_formulas()
        self.assertEqual(result.code, excel_formula.COM_FAILURE)
        self.assertEqual(fatal_reader.call_count, 1)

    def test_transient_error_resets_the_previous_snapshot(self):
        stable = self._selection("2", "=1+1")
        with mock.patch.object(
            excel_formula,
            "read_selected_excel_formulas",
            side_effect=[
                self._success(stable),
                excel_formula.ExcelFormulaResult(excel_formula.SELECTION_CHANGED),
                self._success(stable),
            ],
        ) as reader:
            result = excel_formula.read_stable_selected_excel_formulas()

        self.assertEqual(result.code, excel_formula.SELECTION_CHANGED)
        self.assertIsNone(result.selection)
        self.assertEqual(reader.call_count, 3)


class FormulaTextClipboardTests(unittest.TestCase):
    def test_writer_replaces_clipboard_with_text_and_generated_marker_only(self):
        fake_clipboard = mock.Mock()
        fake_clipboard.RegisterClipboardFormat.side_effect = [49152, 49153]
        fake_con = types.SimpleNamespace(
            CF_TEXT=1,
            CF_BITMAP=2,
            CF_METAFILEPICT=3,
            CF_OEMTEXT=7,
            CF_DIB=8,
            CF_UNICODETEXT=13,
            CF_ENHMETAFILE=14,
            CF_DIBV5=17,
        )

        fake_error = type("FakePyWinError", (Exception,), {})
        fake_pywintypes = types.SimpleNamespace(error=fake_error)
        module_name = "tabledown_windows.win_clipboard"
        package = sys.modules["tabledown_windows"]
        original = sys.modules.pop(module_name, None)
        original_attribute = getattr(package, "win_clipboard", None)
        try:
            with mock.patch.dict(
                sys.modules,
                {
                    "win32clipboard": fake_clipboard,
                    "win32con": fake_con,
                    "pywintypes": fake_pywintypes,
                },
            ):
                clipboard = importlib.import_module(module_name)
                fake_clipboard.GetClipboardData.side_effect = (
                    lambda fmt: (
                        "<수식범위 />"
                        if fmt == fake_con.CF_UNICODETEXT
                        else b"Tabledown"
                    )
                )
                clipboard.write_text_only_clipboard("<수식범위 />")

                with mock.patch.object(
                    clipboard, "clipboard_change_count", return_value=12
                ), self.assertRaises(clipboard.ClipboardChangedError):
                    clipboard.write_text_only_clipboard(
                        "<새범위 />",
                        expected_change_count=11,
                    )

                fake_clipboard.GetClipboardData.side_effect = (
                    lambda fmt: (
                        "<다른범위 />"
                        if fmt == fake_con.CF_UNICODETEXT
                        else b"Tabledown"
                    )
                )
                with self.assertRaises(clipboard.ClipboardWriteError):
                    clipboard.write_text_only_clipboard("<수식범위 />")
        finally:
            sys.modules.pop(module_name, None)
            if original is not None:
                sys.modules[module_name] = original
                package.win_clipboard = original
            elif original_attribute is not None:
                package.win_clipboard = original_attribute
            elif hasattr(package, "win_clipboard"):
                del package.win_clipboard

        # The sequence-mismatch path above adds no EmptyClipboard; the separate
        # read-back mismatch reaches a second write and fails closed.
        self.assertEqual(fake_clipboard.EmptyClipboard.call_count, 2)
        self.assertEqual(
            fake_clipboard.SetClipboardData.call_args_list,
            [
                mock.call(fake_con.CF_UNICODETEXT, "<수식범위 />"),
                mock.call(49153, b"Tabledown"),
                mock.call(fake_con.CF_UNICODETEXT, "<수식범위 />"),
                mock.call(49153, b"Tabledown"),
            ],
        )
        self.assertEqual(
            fake_clipboard.GetClipboardData.call_args_list,
            [
                mock.call(fake_con.CF_UNICODETEXT),
                mock.call(49153),
                mock.call(fake_con.CF_UNICODETEXT),
            ],
        )
        fake_clipboard.EnumClipboardFormats.assert_not_called()


class FormulaExportActionPortableTests(unittest.TestCase):
    """Exercise the real export action without a tray, COM, or OS clipboard."""

    @classmethod
    def setUpClass(cls):
        clipboard = types.ModuleType("tabledown_windows.win_clipboard")
        clipboard.ClipboardChangedError = type("ClipboardChangedError", (Exception,), {})
        clipboard.ClipboardWriteError = type("ClipboardWriteError", (Exception,), {})
        for name in (
            "clipboard_change_count",
            "read_clipboard",
            "write_clipboard",
            "write_text_only_clipboard",
        ):
            setattr(clipboard, name, mock.Mock())
        spec = importlib.util.spec_from_file_location(
            "tabledown_windows._export_action_test",
            WINDOWS_ROOT / "tabledown_windows" / "app.py",
        )
        cls.module = importlib.util.module_from_spec(spec)
        with mock.patch.dict(
            sys.modules,
            {
                "pystray": types.ModuleType("pystray"),
                "tabledown_windows.win_clipboard": clipboard,
            },
        ):
            spec.loader.exec_module(cls.module)

    def _app(self, lang="ko"):
        app = object.__new__(self.module.TabledownWindowsApp)
        app.lang = lang
        app._formula_export_lock = threading.Lock()
        app._clipboard_operation_lock = threading.Lock()
        app._stop_watcher = threading.Event()
        app._show_message_box_async = mock.Mock()
        return app

    @staticmethod
    def _selection(*, partial=False, state="done"):
        return excel_formula.ExcelFormulaSelection(
            "Book.xlsx", "Sheet1", "$A$1", 1, 1,
            (
                excel_formula.ExcelSelectionCell(
                    "$A$1", "2", "=1+1", "=1+1",
                    references_complete=not partial,
                    reference_issues=("read_failed",) if partial else (),
                ),
            ),
            calculation_state=state,
        )

    def _run(
        self, app, selection, *, write_error=None, read_error=None,
        stop_when=None, for_ai=False, serialize_error=None,
    ):
        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        def read():
            if read_error is not None:
                raise read_error
            if stop_when == "read":
                app._stop_watcher.set()
            return excel_formula.ExcelFormulaResult(
                excel_formula.SUCCESS if selection is not None else excel_formula.NO_FORMULAS,
                selection,
            )

        def write(*args, **kwargs):
            if write_error is not None:
                raise write_error
            if stop_when == "write":
                app._stop_watcher.set()

        classifier = self.module.formula_copy_notice_key

        def notice(snapshot):
            # The real writer returns only after read-back verification.
            self.assertTrue(writer.called)
            if stop_when == "notice":
                app._stop_watcher.set()
            return classifier(snapshot)

        if stop_when == "before":
            app._stop_watcher.set()
        with (
            mock.patch.object(self.module.threading, "Thread", ImmediateThread),
            mock.patch.object(self.module, "clipboard_change_count", return_value=31),
            mock.patch.object(self.module, "read_stable_selected_excel_formulas", side_effect=read) as reader,
            mock.patch.object(self.module, "write_text_only_clipboard", side_effect=write) as writer,
            mock.patch.object(self.module, "formula_copy_notice_key", side_effect=notice) as classifier_mock,
            mock.patch.object(self.module, "log"),
            contextlib.ExitStack() as stack,
        ):
            if serialize_error is not None:
                stack.enter_context(mock.patch.object(
                    self.module,
                    "formula_selection_to_ai_xml" if for_ai else "formula_selection_to_xml",
                    side_effect=serialize_error,
                ))
            action = (
                app.copy_selected_excel_formulas_for_ai
                if for_ai else app.copy_selected_excel_formulas
            )
            action(None, None)
        self.assertFalse(app._formula_export_lock.locked())
        return reader, writer, classifier_mock

    def test_success_uses_one_notice_after_verified_write(self):
        cases = (
            (False, "done", "formula_export.success"),
            (True, "done", "formula_export.success.partial_references"),
            (False, "pending", "formula_export.success.calculation_incomplete"),
            (True, "calculating", "formula_export.success.partial_references_and_calculation"),
        )
        for lang in SUPPORTED_LANGUAGES:
            for partial, state, key in cases:
                with self.subTest(lang=lang, partial=partial, state=state):
                    app = self._app(lang)
                    snapshot = self._selection(partial=partial, state=state)
                    _, writer, classifier = self._run(app, snapshot)
                    self.assertIn('수식="=1+1"', writer.call_args.args[0])
                    self.assertIn('값="2"', writer.call_args.args[0])
                    self.assertEqual(writer.call_args.kwargs, {"expected_change_count": 31})
                    classifier.assert_called_once_with(snapshot)
                    app._show_message_box_async.assert_called_once_with(
                        t(key, lang), t("help.title", lang)
                    )

    def test_read_failures_never_show_success_notice(self):
        for error in (None, RuntimeError("reader unavailable")):
            with self.subTest(error=type(error).__name__):
                app = self._app()
                _, writer, classifier = self._run(app, None, read_error=error)
                writer.assert_not_called()
                classifier.assert_not_called()
                code = "export_failed" if error else "no_formulas"
                app._show_message_box_async.assert_called_once_with(
                    t(f"formula_export.error.{code}", app.lang), t("help.title", app.lang)
                )

    def test_write_failure_or_new_copy_never_show_success_notice(self):
        for error_type, code in (
            (self.module.ClipboardWriteError, "clipboard_write_failed"),
            (self.module.ClipboardChangedError, "clipboard_changed"),
        ):
            with self.subTest(code=code):
                app = self._app()
                _, writer, classifier = self._run(
                    app, self._selection(partial=True), write_error=error_type(code)
                )
                writer.assert_called_once()
                classifier.assert_not_called()
                app._show_message_box_async.assert_called_once_with(
                    t(f"formula_export.error.{code}", app.lang), t("help.title", app.lang)
                )

    def test_quit_never_shows_success_notice(self):
        for stop_when in ("before", "read", "write", "notice"):
            with self.subTest(stop_when=stop_when):
                app = self._app()
                reader, writer, classifier = self._run(
                    app, self._selection(partial=True), stop_when=stop_when
                )
                if stop_when == "before":
                    reader.assert_not_called()
                if stop_when in ("before", "read"):
                    writer.assert_not_called()
                if stop_when != "notice":
                    classifier.assert_not_called()
                app._show_message_box_async.assert_not_called()

    def test_failed_export_can_retry_with_a_single_success_notice(self):
        app = self._app()
        snapshot = self._selection(partial=True)
        self._run(app, snapshot, write_error=self.module.ClipboardWriteError("failed"))
        app._show_message_box_async.reset_mock()
        _, writer, classifier = self._run(app, snapshot)
        writer.assert_called_once()
        classifier.assert_called_once_with(snapshot)
        app._show_message_box_async.assert_called_once_with(
            t("formula_export.success.partial_references", app.lang), t("help.title", app.lang)
        )

    def test_ai_and_original_actions_use_their_serializer_and_one_stable_read(self):
        for for_ai in (False, True):
            with self.subTest(for_ai=for_ai):
                app = self._app()
                snapshot = self._selection()
                with (
                    mock.patch.object(self.module, "formula_selection_to_xml", return_value="original") as original,
                    mock.patch.object(self.module, "formula_selection_to_ai_xml", return_value="compact") as compact,
                ):
                    reader, writer, classifier = self._run(app, snapshot, for_ai=for_ai)
                reader.assert_called_once_with()
                selected, unused = (compact, original) if for_ai else (original, compact)
                selected.assert_called_once_with(snapshot)
                unused.assert_not_called()
                writer.assert_called_once_with(
                    "compact" if for_ai else "original", expected_change_count=31,
                )
                classifier.assert_called_once_with(snapshot)

    def test_ai_success_preserves_partial_reference_and_calculation_notices(self):
        cases = (
            (False, "done", "formula_export.success.ai"),
            (True, "done", "formula_export.success.partial_references"),
            (False, "pending", "formula_export.success.calculation_incomplete"),
            (True, "calculating", "formula_export.success.partial_references_and_calculation"),
        )
        for lang in SUPPORTED_LANGUAGES:
            for partial, state, key in cases:
                with self.subTest(lang=lang, partial=partial, state=state):
                    app = self._app(lang)
                    snapshot = self._selection(partial=partial, state=state)
                    reader, writer, classifier = self._run(app, snapshot, for_ai=True)
                    reader.assert_called_once_with()
                    writer.assert_called_once()
                    classifier.assert_called_once_with(snapshot)
                    app._show_message_box_async.assert_called_once_with(
                        t(key, lang), t("help.title", lang),
                    )

    def test_ai_failures_release_gate_and_allow_manual_retry_without_false_success(self):
        cases = (
            ("no_formulas", None, {}, False),
            ("export_failed", None, {"read_error": RuntimeError("read unavailable")}, False),
            ("export_failed", self._selection(), {"serialize_error": ValueError("invalid")}, False),
            ("output_too_large", self._selection(), {"serialize_error": self.module.FormulaXmlTooLargeError("limit")}, False),
            ("clipboard_write_failed", self._selection(), {"write_error": self.module.ClipboardWriteError("failed")}, True),
            ("clipboard_changed", self._selection(), {"write_error": self.module.ClipboardChangedError("changed")}, True),
        )
        for code, snapshot, kwargs, tried_write in cases:
            with self.subTest(code=code, kwargs=tuple(kwargs)):
                app = self._app()
                _, writer, classifier = self._run(app, snapshot, for_ai=True, **kwargs)
                self.assertEqual(writer.call_count, int(tried_write))
                classifier.assert_not_called()
                app._show_message_box_async.assert_called_once_with(
                    t(f"formula_export.error.{code}", app.lang), t("help.title", app.lang),
                )
                app._show_message_box_async.reset_mock()
                _, writer, classifier = self._run(app, self._selection(), for_ai=True)
                writer.assert_called_once()
                classifier.assert_called_once()
                app._show_message_box_async.assert_called_once_with(
                    t("formula_export.success.ai", app.lang), t("help.title", app.lang),
                )

    def test_ai_quit_skips_pending_write_and_success(self):
        for stop_when in ("before", "read", "write", "notice"):
            with self.subTest(stop_when=stop_when):
                app = self._app()
                reader, writer, classifier = self._run(
                    app, self._selection(), for_ai=True, stop_when=stop_when,
                )
                if stop_when == "before":
                    reader.assert_not_called()
                if stop_when in ("before", "read"):
                    writer.assert_not_called()
                if stop_when != "notice":
                    classifier.assert_not_called()
                app._show_message_box_async.assert_not_called()

    def test_ai_and_original_actions_share_gate_and_pause_watcher(self):
        for ai_first in (False, True):
            with self.subTest(ai_first=ai_first):
                app = self._app()
                workers = []

                class HeldThread:
                    def __init__(self, *, target, name, daemon):
                        self.target = target
                        workers.append(self)

                    def start(self):
                        pass

                first, second = (
                    (app.copy_selected_excel_formulas_for_ai, app.copy_selected_excel_formulas)
                    if ai_first else
                    (app.copy_selected_excel_formulas, app.copy_selected_excel_formulas_for_ai)
                )
                with (
                    mock.patch.object(self.module.threading, "Thread", HeldThread),
                    mock.patch.object(self.module, "clipboard_change_count", return_value=31) as count,
                    mock.patch.object(self.module, "read_clipboard") as read_clipboard,
                    mock.patch.object(self.module, "write_clipboard") as write_clipboard,
                    mock.patch.object(self.module, "read_stable_selected_excel_formulas") as read_excel,
                    mock.patch.object(self.module, "log"),
                ):
                    first(None, None)
                    count.reset_mock()
                    second(None, None)
                    self.assertEqual(len(workers), 1)
                    self.assertFalse(app._augment_clipboard())
                    count.assert_not_called()
                    read_clipboard.assert_not_called()
                    write_clipboard.assert_not_called()
                    read_excel.assert_not_called()
                    app._show_message_box_async.assert_called_once_with(
                        t("formula_export.error.in_progress", app.lang), t("help.title", app.lang),
                    )
                    app._stop_watcher.set()
                    workers[0].target()
                    self.assertFalse(app._formula_export_lock.locked())
                    app._stop_watcher.clear()
                    second(None, None)
                    self.assertEqual(len(workers), 2)
                    app._stop_watcher.set()
                    workers[1].target()
                    self.assertFalse(app._formula_export_lock.locked())

    def test_ai_startup_failures_release_gate_for_retry(self):
        class FailingThread:
            def __init__(self, *, target, name, daemon):
                pass

            def start(self):
                raise RuntimeError("thread unavailable")

        for fail_at in ("clipboard_count", "thread"):
            with self.subTest(fail_at=fail_at):
                app = self._app()
                with (
                    mock.patch.object(self.module, "clipboard_change_count", return_value=31) as count,
                    mock.patch.object(self.module.threading, "Thread", FailingThread),
                    mock.patch.object(self.module, "read_stable_selected_excel_formulas") as reader,
                    mock.patch.object(self.module, "write_text_only_clipboard") as writer,
                    mock.patch.object(self.module, "log"),
                ):
                    if fail_at == "clipboard_count":
                        count.side_effect = RuntimeError("count unavailable")
                    app.copy_selected_excel_formulas_for_ai(None, None)
                self.assertFalse(app._formula_export_lock.locked())
                reader.assert_not_called()
                writer.assert_not_called()
                app._show_message_box_async.assert_called_once_with(
                    t("formula_export.error.export_failed", app.lang), t("help.title", app.lang),
                )
                app._show_message_box_async.reset_mock()
                _, writer, _ = self._run(app, self._selection(), for_ai=True)
                writer.assert_called_once()
                app._show_message_box_async.assert_called_once_with(
                    t("formula_export.success.ai", app.lang), t("help.title", app.lang),
                )

    def test_ai_menu_is_additive_localized_and_keeps_original_hotkey(self):
        class Menu(list):
            SEPARATOR = object()

            def __init__(self, *items):
                super().__init__(items)

        def menu_item(text, action, **kwargs):
            return types.SimpleNamespace(text=text, action=action)

        for lang in SUPPORTED_LANGUAGES:
            with self.subTest(lang=lang):
                app = self._app(lang)
                app.login_supported = False
                fake_pystray = types.SimpleNamespace(Menu=Menu, MenuItem=menu_item)
                with mock.patch.object(self.module, "pystray", fake_pystray):
                    menu = app._build_menu()
                actions = {item.text: item.action for item in menu if item is not Menu.SEPARATOR}
                self.assertEqual(actions[t("menu.copy_excel_formulas", lang)], app.copy_selected_excel_formulas)
                self.assertEqual(actions[t("menu.copy_excel_formulas_ai", lang)], app.copy_selected_excel_formulas_for_ai)
                self.assertEqual(t("menu.copy_excel_formulas_ai", lang), "AI용 간결 복사" if lang == "ko" else "Copy compact XML for AI")
                for key in ("formula_export.success.ai", "help.message"):
                    self.assertNotEqual(t(key, lang), key)
                self.assertIn("추정" if lang == "ko" else "inferred", t("help.message", lang))
                self.assertIn("제목·항목명" if lang == "ko" else "headers and item names", t("help.message", lang))
                with (
                    mock.patch.object(app, "copy_selected_excel_formulas") as original,
                    mock.patch.object(app, "copy_selected_excel_formulas_for_ai") as compact,
                ):
                    app._copy_selected_excel_formulas_from_hotkey()
                original.assert_called_once_with(None, None)
                compact.assert_not_called()


class _FakeWinrtTask:
    """Stand-in for a WinRT StartupTask object.

    ``state`` returns a real ``StartupTaskState`` member so the production
    ``_state_name`` mapping is exercised. ``request_enable_async`` returns a
    coroutine (awaitable, like the real ``IAsyncOperation``); ``disable`` is
    synchronous, matching the real API.
    """

    def __init__(self, state, becomes=None):
        self._state = state
        self._becomes = becomes
        self.enable_requested = False
        self.disabled = False

    @property
    def state(self):
        return self._state

    def request_enable_async(self):
        async def _op():
            self.enable_requested = True
            if self._becomes is not None:
                self._state = self._becomes
        return _op()

    def disable(self) -> None:
        self.disabled = True
        if self._becomes is not None:
            self._state = self._becomes


class _FakeStartupTaskApi:
    """Stand-in for the winsdk ``StartupTask`` class (its ``get_async``)."""

    def __init__(self, task: _FakeWinrtTask):
        self._task = task

    def get_async(self, _task_id):
        async def _op():
            return self._task
        return _op()


@contextlib.contextmanager
def _patched_api(task: _FakeWinrtTask):
    # Swap winsdk's StartupTask for a fake; the real _run/asyncio/_state_name
    # orchestration runs against it on the production code path.
    with mock.patch.object(startup_task, "StartupTask", _FakeStartupTaskApi(task)):
        yield


class StartupTaskDegradeTests(unittest.TestCase):
    def test_degrades_without_winrt(self):
        # No winsdk / no package identity: every entry point is safe and falsey,
        # so the menu just omits the toggle (mirrors macOS login_item).
        with mock.patch.object(startup_task, "StartupTask", None):
            self.assertIsNone(startup_task.current_state())
            self.assertFalse(startup_task.is_supported())
            self.assertFalse(startup_task.is_enabled())
            self.assertEqual(startup_task.set_enabled(True), "unavailable")


@unittest.skipUnless(
    startup_task.StartupTaskState is not None, "winsdk StartupTaskState required"
)
class StartupTaskTests(unittest.TestCase):
    def _state(self, name):
        return getattr(startup_task.StartupTaskState, name)

    def test_current_state_and_is_enabled(self):
        with _patched_api(_FakeWinrtTask(self._state("ENABLED"))):
            self.assertEqual(startup_task.current_state(), "enabled")
            self.assertTrue(startup_task.is_supported())
            self.assertTrue(startup_task.is_enabled())

    def test_enable_requests_and_reports_enabled(self):
        fake = _FakeWinrtTask(self._state("DISABLED"), becomes=self._state("ENABLED"))
        with _patched_api(fake):
            self.assertEqual(startup_task.set_enabled(True), "enabled")
        self.assertTrue(fake.enable_requested)

    def test_disable_calls_disable(self):
        fake = _FakeWinrtTask(self._state("ENABLED"), becomes=self._state("DISABLED"))
        with _patched_api(fake):
            self.assertEqual(startup_task.set_enabled(False), "disabled")
        self.assertTrue(fake.disabled)

    def test_blocked_by_user_is_not_an_enabled_state(self):
        # Windows can keep an enable request DISABLED_BY_USER; the read-back name
        # must surface that and never count as enabled.
        fake = _FakeWinrtTask(self._state("DISABLED_BY_USER"))  # stays put
        with _patched_api(fake):
            status = startup_task.set_enabled(True)
        self.assertTrue(fake.enable_requested)
        self.assertEqual(status, "disabled_by_user")
        self.assertNotIn(status, startup_task.ENABLED_STATES)


@unittest.skipUnless(sys.platform.startswith("win"), "tray app is Windows-only")
class LoginMenuTests(unittest.TestCase):
    def _make_app(self, *, supported: bool, enabled: bool = False):
        # app.__init__ seeds login_supported/login_enabled from one
        # current_state() read; None means "unsupported, hide the toggle".
        state = ("enabled" if enabled else "disabled") if supported else None
        with mock.patch.object(startup_task, "current_state", return_value=state):
            from tabledown_windows.app import TabledownWindowsApp

            return TabledownWindowsApp()

    def _labels(self, app):
        return [getattr(item, "text", "") for item in app.icon.menu]

    def test_menu_includes_login_item_when_supported(self):
        app = self._make_app(supported=True)
        self.assertIn(t("menu.login_item", app.lang), self._labels(app))

    def test_menu_omits_login_item_when_unsupported(self):
        app = self._make_app(supported=False)
        self.assertNotIn(t("menu.login_item", app.lang), self._labels(app))

    def test_toggle_blocked_by_user_stays_unchecked_and_warns(self):
        app = self._make_app(supported=True, enabled=False)
        shown = []
        with mock.patch.object(startup_task, "set_enabled", return_value="disabled_by_user"), \
             mock.patch.object(app, "_refresh_menu"), \
             mock.patch.object(
                 app, "_show_message_box_async", side_effect=lambda msg, _title: shown.append(msg)
             ):
            app.toggle_login_item(None, None)
        self.assertFalse(app.login_enabled)  # checkmark stays truthful
        self.assertEqual(len(shown), 1)      # user is told why

    def test_toggle_enable_checks_and_is_silent(self):
        app = self._make_app(supported=True, enabled=False)
        shown = []
        with mock.patch.object(startup_task, "set_enabled", return_value="enabled"), \
             mock.patch.object(app, "_refresh_menu"), \
             mock.patch.object(
                 app, "_show_message_box_async", side_effect=lambda msg, _title: shown.append(msg)
             ):
            app.toggle_login_item(None, None)
        self.assertTrue(app.login_enabled)
        self.assertEqual(shown, [])  # success is silent — the checkmark says it

    def test_menu_includes_fill_blanks_toggle(self):
        app = self._make_app(supported=False)
        self.assertIn(t("menu.fill_blanks", app.lang), self._labels(app))

    def test_menu_includes_excel_formula_export(self):
        app = self._make_app(supported=False)
        self.assertIn(t("menu.copy_excel_formulas", app.lang), self._labels(app))

    def test_excel_formula_hotkey_uses_ctrl_alt_e_and_menu_action(self):
        app = self._make_app(supported=False)

        self.assertEqual(
            app._formula_hotkey._modifiers,
            hotkey.MOD_CONTROL | hotkey.MOD_ALT | hotkey.MOD_NOREPEAT,
        )
        self.assertEqual(app._formula_hotkey._vk, hotkey.VK_E)
        with mock.patch.object(app, "copy_selected_excel_formulas") as export:
            app._formula_hotkey._callback()
        export.assert_called_once_with(None, None)

    def test_run_starts_toggle_and_formula_hotkeys(self):
        app = self._make_app(supported=False)

        with mock.patch("tabledown_windows.app.threading.Thread") as worker, \
             mock.patch.object(app._hotkey, "start", return_value=True) as toggle_start, \
             mock.patch.object(
                 app._formula_hotkey, "start", return_value=True
             ) as formula_start, \
             mock.patch.object(app.icon, "run") as icon_run:
            app.run()

        worker.return_value.start.assert_called_once_with()
        toggle_start.assert_called_once_with()
        formula_start.assert_called_once_with()
        icon_run.assert_called_once_with(setup=app._on_ready)

    def test_quit_stops_toggle_and_formula_hotkeys(self):
        app = self._make_app(supported=False)

        with mock.patch.object(app._hotkey, "stop") as toggle_stop, \
             mock.patch.object(app._formula_hotkey, "stop") as formula_stop, \
             mock.patch.object(app.icon, "stop") as icon_stop:
            app.quit_app(None, None)

        self.assertTrue(app._stop_watcher.is_set())
        toggle_stop.assert_called_once_with()
        formula_stop.assert_called_once_with()
        icon_stop.assert_called_once_with()

    def test_excel_formula_export_starts_daemon_worker(self):
        app = self._make_app(supported=False)
        created = []

        class FakeThread:
            def __init__(self, *, target, name, daemon):
                created.append((target, name, daemon))

            def start(self):
                return None

        with mock.patch("tabledown_windows.app.threading.Thread", FakeThread):
            app.copy_selected_excel_formulas(None, None)

        self.assertEqual(len(created), 1)
        self.assertEqual(created[0][1], "TabledownWindowsExcelFormulaExport")
        self.assertTrue(created[0][2])

    def test_excel_formula_export_ignores_click_while_worker_is_active(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        created = []
        result = types.SimpleNamespace(ok=False, code=excel_formula.NO_FORMULAS)

        class HeldThread:
            def __init__(self, *, target, name, daemon):
                self.target = target
                created.append(self)

            def start(self):
                return None

        with mock.patch.object(app_module.threading, "Thread", HeldThread), \
             mock.patch.object(
                 app_module, "read_stable_selected_excel_formulas", return_value=result
             ), \
             mock.patch.object(app, "_show_message_box_async") as show:
            app.copy_selected_excel_formulas(None, None)
            app.copy_selected_excel_formulas(None, None)
            self.assertEqual(len(created), 1)
            show.assert_called_once_with(
                t("formula_export.error.in_progress", app.lang),
                t("help.title", app.lang),
            )

            # Finishing the first worker releases the gate; a later click can
            # now create a fresh export worker.
            created[0].target()
            app.copy_selected_excel_formulas(None, None)
            self.assertEqual(len(created), 2)

    def test_excel_formula_thread_start_failure_releases_gate(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)

        class FailingThread:
            def __init__(self, *, target, name, daemon):
                return None

            def start(self):
                raise OSError("thread unavailable")

        with mock.patch.object(app_module.threading, "Thread", FailingThread), \
             mock.patch.object(app, "_show_message_box_async"):
            app.copy_selected_excel_formulas(None, None)

        self.assertTrue(app._formula_export_lock.acquire(blocking=False))
        app._formula_export_lock.release()

    def test_excel_formula_export_writes_xml_and_shows_success(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        selection = StableExcelFormulaReaderTests._selection("2", "=1+1")
        result = types.SimpleNamespace(ok=True, selection=selection)
        shown = []

        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(app_module.threading, "Thread", ImmediateThread), \
             mock.patch.object(app_module, "clipboard_change_count", return_value=31), \
             mock.patch.object(app_module, "read_stable_selected_excel_formulas", return_value=result), \
             mock.patch.object(app_module, "formula_selection_to_xml", return_value="<수식범위 />") as serialize, \
             mock.patch.object(app_module, "write_text_only_clipboard") as write, \
             mock.patch.object(
                 app,
                 "_show_message_box_async",
                 side_effect=lambda message, _title: shown.append(message),
             ):
            app.copy_selected_excel_formulas(None, None)

        serialize.assert_called_once_with(selection)
        write.assert_called_once_with(
            "<수식범위 />",
            expected_change_count=31,
        )
        self.assertEqual(shown, [t("formula_export.success", app.lang)])

    def test_excel_formula_export_reports_output_size_without_writing(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        result = types.SimpleNamespace(ok=True, selection=object())
        shown = []

        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(app_module.threading, "Thread", ImmediateThread), \
             mock.patch.object(app_module, "clipboard_change_count", return_value=31), \
             mock.patch.object(
                 app_module,
                 "read_stable_selected_excel_formulas",
                 return_value=result,
             ), \
             mock.patch.object(
                 app_module,
                 "formula_selection_to_xml",
                 side_effect=app_module.FormulaXmlTooLargeError("too large"),
             ), \
             mock.patch.object(app_module, "write_text_only_clipboard") as write, \
             mock.patch.object(
                 app,
                 "_show_message_box_async",
                 side_effect=lambda message, _title: shown.append(message),
             ):
            app.copy_selected_excel_formulas(None, None)

        write.assert_not_called()
        self.assertEqual(
            shown,
            [t("formula_export.error.output_too_large", app.lang)],
        )

    def test_excel_formula_export_reports_verified_write_failure(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        result = types.SimpleNamespace(ok=True, selection=object())
        shown = []

        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(app_module.threading, "Thread", ImmediateThread), \
             mock.patch.object(app_module, "clipboard_change_count", return_value=31), \
             mock.patch.object(
                 app_module,
                 "read_stable_selected_excel_formulas",
                 return_value=result,
             ), \
             mock.patch.object(
                 app_module,
                 "formula_selection_to_xml",
                 return_value="<수식범위 />",
             ), \
             mock.patch.object(
                 app_module,
                 "write_text_only_clipboard",
                 side_effect=app_module.ClipboardWriteError("write failed"),
             ), \
             mock.patch.object(
                 app,
                 "_show_message_box_async",
                 side_effect=lambda message, _title: shown.append(message),
             ):
            app.copy_selected_excel_formulas(None, None)

        self.assertEqual(
            shown,
            [t("formula_export.error.clipboard_write_failed", app.lang)],
        )

    def test_excel_formula_export_shows_localized_adapter_error(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        result = types.SimpleNamespace(ok=False, code=excel_formula.NO_FORMULAS)
        shown = []

        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(app_module.threading, "Thread", ImmediateThread), \
             mock.patch.object(app_module, "clipboard_change_count", return_value=31), \
             mock.patch.object(app_module, "read_stable_selected_excel_formulas", return_value=result), \
             mock.patch.object(app_module, "write_text_only_clipboard") as write, \
             mock.patch.object(
                 app,
                 "_show_message_box_async",
                 side_effect=lambda message, _title: shown.append(message),
             ):
            app.copy_selected_excel_formulas(None, None)

        write.assert_not_called()
        self.assertEqual(
            shown, [t("formula_export.error.no_formulas", app.lang)]
        )

    def test_excel_formula_export_preserves_newer_clipboard(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        result = types.SimpleNamespace(ok=True, selection=object())
        shown = []

        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(app_module.threading, "Thread", ImmediateThread), \
             mock.patch.object(app_module, "clipboard_change_count", return_value=31), \
             mock.patch.object(
                 app_module, "read_stable_selected_excel_formulas", return_value=result
             ), \
             mock.patch.object(
                 app_module, "formula_selection_to_xml", return_value="<수식범위 />"
             ), \
             mock.patch.object(
                 app_module,
                 "write_text_only_clipboard",
                 side_effect=app_module.ClipboardChangedError(
                     "clipboard_changed"
                 ),
             ), \
             mock.patch.object(
                 app,
                 "_show_message_box_async",
                 side_effect=lambda message, _title: shown.append(message),
             ):
            app.copy_selected_excel_formulas(None, None)

        self.assertEqual(
            shown,
            [t("formula_export.error.clipboard_changed", app.lang)],
        )

    def test_watcher_rechecks_export_gate_inside_clipboard_lock(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)

        class ExportStartsOnEnter:
            def __enter__(inner_self):
                self.assertTrue(
                    app._formula_export_lock.acquire(blocking=False)
                )
                return inner_self

            def __exit__(inner_self, exc_type, exc_value, traceback):
                app._formula_export_lock.release()

        app._clipboard_operation_lock = ExportStartsOnEnter()
        with mock.patch.object(app_module, "clipboard_change_count") as count, \
             mock.patch.object(app_module, "read_clipboard") as read, \
             mock.patch.object(app_module, "write_clipboard") as write:
            processed = app._augment_clipboard()

        self.assertFalse(processed)
        count.assert_not_called()
        read.assert_not_called()
        write.assert_not_called()

    def test_watcher_keeps_generation_pending_after_clipboard_mismatch(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        app.enabled = True
        app._last_change_count = 10

        def stop_after_sleep(_seconds):
            app._stop_watcher.set()

        with mock.patch.object(
            app_module.time,
            "sleep",
            side_effect=stop_after_sleep,
        ), mock.patch.object(
            app_module,
            "clipboard_change_count",
            return_value=11,
        ), mock.patch.object(
            app_module,
            "read_clipboard",
            return_value={"text": "source"},
        ), mock.patch.object(
            app_module,
            "converted_clipboard",
            return_value={"text": "updated"},
        ), mock.patch.object(
            app_module,
            "write_clipboard",
            side_effect=app_module.ClipboardChangedError("clipboard changed"),
        ):
            app._watch_clipboard()

        self.assertEqual(app._last_change_count, 10)

    def test_excel_formula_export_after_quit_skips_excel_and_clipboard(self):
        from tabledown_windows import app as app_module

        app = self._make_app(supported=False)
        app._stop_watcher.set()

        class ImmediateThread:
            def __init__(self, *, target, name, daemon):
                self.target = target

            def start(self):
                self.target()

        with mock.patch.object(app_module.threading, "Thread", ImmediateThread), \
             mock.patch.object(app_module, "clipboard_change_count", return_value=31), \
             mock.patch.object(app_module, "read_stable_selected_excel_formulas") as read, \
             mock.patch.object(app_module, "write_text_only_clipboard") as write, \
             mock.patch.object(app, "_show_message_box_async") as show:
            app.copy_selected_excel_formulas(None, None)

        read.assert_not_called()
        write.assert_not_called()
        show.assert_not_called()
        self.assertTrue(app._formula_export_lock.acquire(blocking=False))
        app._formula_export_lock.release()

    def test_toggle_fill_blanks_flips_and_persists(self):
        app = self._make_app(supported=False)
        self.assertFalse(app.fill_blanks)  # off by default
        saved = []
        from tabledown_windows.app import FILL_BLANKS_KEY

        with mock.patch("tabledown_windows.app.save_setting",
                        side_effect=lambda key, value: saved.append((key, value))), \
             mock.patch.object(app, "_refresh_menu"):
            app.toggle_fill_blanks(None, None)
        self.assertTrue(app.fill_blanks)
        self.assertEqual(saved, [(FILL_BLANKS_KEY, True)])


class SingleInstanceTests(unittest.TestCase):
    """The named-mutex guard that stops a second tray (and second watcher)."""

    def setUp(self):
        # acquire_single_instance stashes the live handle in a module global;
        # clear it so one test's claim can't leak into the next.
        single_instance._held_handle = None
        self.addCleanup(setattr, single_instance, "_held_handle", None)

    def test_first_instance_acquires_and_holds_handle(self):
        # Fresh mutex (not pre-existing) -> this is the owner; the handle is kept
        # alive for the process lifetime so the mutex outlives the call.
        with mock.patch.object(single_instance, "_create_mutex", return_value=(4321, False)):
            self.assertTrue(single_instance.acquire_single_instance())
        self.assertEqual(single_instance._held_handle, 4321)

    def test_second_instance_is_blocked(self):
        # CreateMutexW reported ERROR_ALREADY_EXISTS -> a sibling owns it; bail
        # out and do not retain a handle.
        with mock.patch.object(single_instance, "_create_mutex", return_value=(4321, True)):
            self.assertFalse(single_instance.acquire_single_instance())
        self.assertIsNone(single_instance._held_handle)

    def test_missing_kernel_degrades_to_allowed(self):
        # Non-Windows host (no kernel32): _create_mutex yields (None, False), so
        # the app is allowed to start rather than refusing over a missing guard.
        with mock.patch.object(single_instance, "_create_mutex", return_value=(None, False)):
            self.assertTrue(single_instance.acquire_single_instance())


@unittest.skipUnless(sys.platform.startswith("win"), "tray app is Windows-only")
class SingleInstanceMainTests(unittest.TestCase):
    def test_main_skips_tray_when_already_running(self):
        # The guard must short-circuit main() before the tray app is built, so a
        # second launch adds neither an icon nor a clipboard watcher.
        from tabledown_windows import app

        with mock.patch.object(app.single_instance, "acquire_single_instance", return_value=False), \
             mock.patch.object(app, "TabledownWindowsApp") as fake_app:
            app.main()
        fake_app.assert_not_called()

    def test_main_builds_tray_when_first(self):
        from tabledown_windows import app

        with mock.patch.object(app.single_instance, "acquire_single_instance", return_value=True), \
             mock.patch.object(app, "TabledownWindowsApp") as fake_app:
            app.main()
        fake_app.assert_called_once_with()
        fake_app.return_value.run.assert_called_once_with()


class HotkeyTests(unittest.TestCase):
    """The Ctrl+Alt+T global hotkey (RegisterHotKey on a private pump thread).

    No real user32 is touched — the worker thread/message loop is Windows-only,
    so we exercise the register/dispatch/degrade logic with stubs (mirrors the
    single-instance tests, which never spawn a real second process).
    """

    def _hk(self, callback=None):
        return hotkey.GlobalHotkey(
            hotkey.MOD_CONTROL | hotkey.MOD_ALT | hotkey.MOD_NOREPEAT,
            hotkey.VK_T,
            callback or (lambda: None),
        )

    def test_start_degrades_without_user32(self):
        # Non-Windows host (or no user32): start() reports False and spawns no
        # thread, so the tray runs fine without the accelerator.
        hk = self._hk()
        with mock.patch.object(hotkey, "_load_user32", return_value=None):
            self.assertFalse(hk.start())
        self.assertFalse(hk.registered)
        self.assertIsNone(hk._thread)

    def test_register_calls_registerhotkey_with_combo(self):
        hk = self._hk()
        user32 = mock.Mock()
        user32.RegisterHotKey.return_value = 1
        self.assertTrue(hk._register(user32))
        user32.RegisterHotKey.assert_called_once_with(
            None,
            hotkey._HOTKEY_ID,
            hotkey.MOD_CONTROL | hotkey.MOD_ALT | hotkey.MOD_NOREPEAT,
            hotkey.VK_T,
        )

    def test_formula_hotkey_registers_ctrl_alt_e(self):
        hk = hotkey.GlobalHotkey(
            hotkey.MOD_CONTROL | hotkey.MOD_ALT | hotkey.MOD_NOREPEAT,
            hotkey.VK_E,
            lambda: None,
        )
        user32 = mock.Mock()
        user32.RegisterHotKey.return_value = 1

        self.assertTrue(hk._register(user32))
        user32.RegisterHotKey.assert_called_once_with(
            None,
            hotkey._HOTKEY_ID,
            hotkey.MOD_CONTROL | hotkey.MOD_ALT | hotkey.MOD_NOREPEAT,
            0x45,
        )

    def test_register_fails_when_combo_busy(self):
        # RegisterHotKey returns 0 when another app already owns the combo.
        hk = self._hk()
        user32 = mock.Mock()
        user32.RegisterHotKey.return_value = 0
        self.assertFalse(hk._register(user32))

    def test_handle_message_fires_on_matching_hotkey(self):
        fired = []
        hk = self._hk(lambda: fired.append(True))
        hk._handle_message(mock.Mock(message=hotkey._WM_HOTKEY, wParam=hotkey._HOTKEY_ID))
        self.assertEqual(fired, [True])

    def test_handle_message_ignores_other_messages(self):
        fired = []
        hk = self._hk(lambda: fired.append(True))
        hk._handle_message(mock.Mock(message=0x0100, wParam=hotkey._HOTKEY_ID))  # not WM_HOTKEY
        hk._handle_message(mock.Mock(message=hotkey._WM_HOTKEY, wParam=999))      # other hotkey id
        self.assertEqual(fired, [])

    def test_stop_before_start_is_safe(self):
        hk = self._hk()
        hk.stop()  # no thread yet — must be a quiet no-op, not raise


if __name__ == "__main__":
    unittest.main()
