from __future__ import annotations

from pathlib import Path
import plistlib
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, call, patch

from Foundation import NSAppleScript

from tablemark import clipboard
from tablemark.converter.formula_export import ExcelFormulaSelection, ExcelSelectionCell
from tablemark.excel_formula import (
    AREA_RESULT_STATUS,
    AUTOMATION_DENIED,
    EXCEL_MASK_SCRIPT,
    EXCEL_NOT_RUNNING,
    EXECUTION_FAILED,
    INVALID_RESPONSE,
    MASK_RESULT_STATUS,
    MAX_FORMULA_CHUNK_CELLS,
    MULTIPLE_AREAS,
    NO_FORMULAS,
    NO_SELECTION,
    SELECTION_CHANGED,
    TOO_FRAGMENTED,
    TOO_MANY_CELLS,
    TOO_MUCH_TEXT,
    ExcelFormulaError,
    NSAppleScriptExecutor,
    _build_formula_read_script,
    _parse_mask_result,
    _rectangle_bounds,
    parse_excel_formula_result,
    read_selected_excel_formulas,
    read_stable_selected_excel_formulas,
)
from tablemark.i18n import t


ROOT = Path(__file__).resolve().parents[1]


def mask_payload(
    *,
    workbook="Book.xlsx",
    sheet="Sheet1",
    address="$A$1",
    rows=1,
    columns=1,
    mask=None,
):
    flags = ["true"] if mask is None else mask
    return [
        MASK_RESULT_STATUS,
        workbook,
        sheet,
        address,
        str(rows),
        str(columns),
        flags,
    ]


def area_payload(
    areas,
    *,
    workbook="Book.xlsx",
    sheet="Sheet1",
    address="$A$1",
    value_areas=None,
):
    start_row, start_column, end_row, end_column = _rectangle_bounds(address)
    rows = end_row - start_row + 1
    columns = end_column - start_column + 1
    if value_areas is None:
        value_areas = [
            [
                address,
                str(rows),
                str(columns),
                [["value", f"value-{index + 1}"] for index in range(rows * columns)],
            ]
        ]
    return [
        AREA_RESULT_STATUS,
        workbook,
        sheet,
        address,
        str(rows),
        str(columns),
        value_areas,
        areas,
    ]


def value_area(values, *, address="$A$1", rows=1, columns=1):
    return [
        address,
        str(rows),
        str(columns),
        [
            ["blank", ""] if value is None else ["value", value]
            for value in values
        ],
    ]


def single_area(
    formula_a1="=1+1",
    formula_r1c1="=1+1",
    *,
    address="$A$1",
):
    return [
        address,
        "1",
        "1",
        formula_a1,
        "present" if formula_r1c1 is not None else "missing",
        formula_r1c1 or "",
    ]


class FakeExecutor:
    def __init__(self, payloads=(), *, running=True, error=None):
        self.payloads = list(payloads)
        self.running = running
        self.error = error
        self.sources = []

    def is_excel_running(self):
        return self.running

    def run(self, source):
        self.sources.append(source)
        if self.error is not None:
            raise self.error
        if not self.payloads:
            raise AssertionError("unexpected executor call")
        return self.payloads.pop(0)


class ExcelFormulaReaderTests(unittest.TestCase):
    @unittest.skipUnless(
        Path("/Applications/Microsoft Excel.app").exists(),
        "Microsoft Excel terminology is required to compile the scripts",
    )
    def test_both_generated_applescripts_compile(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$CV$100",
                rows=100,
                columns=100,
                mask="1" * 10_000,
            )
        )

        for source in (EXCEL_MASK_SCRIPT, _build_formula_read_script(plan)):
            with self.subTest(source_length=len(source)):
                script = NSAppleScript.alloc().initWithSource_(source)
                compiled, error_info = script.compileAndReturnError_(None)
                self.assertTrue(compiled, error_info)

    def test_two_stage_reader_fetches_all_values_and_only_formula_rectangles(self):
        formula = '=LET(값,A1,"tab:\tline:\n"&값)'
        workbook = 'Q"\\\nBook.xlsx'
        sheet = "요약\t검토"
        first = mask_payload(
            workbook=workbook,
            sheet=sheet,
            address="$B$2:$C$3",
            rows=2,
            columns=2,
            mask=["true", "false", "false", "true"],
        )
        second = area_payload(
            [
                single_area(
                    formula,
                    '=LET(값,RC[-1],"tab:\tline:\n"&값)',
                    address="$B$2",
                ),
                single_area("=B2*2", "=R[-1]C[-1]*2", address="$C$3"),
            ],
            workbook=workbook,
            sheet=sheet,
            address="$B$2:$C$3",
        )
        executor = FakeExecutor([first, second])

        selection = read_selected_excel_formulas(executor)

        self.assertEqual(
            [cell.address for cell in selection.cells],
            ["$B$2", "$C$2", "$B$3", "$C$3"],
        )
        self.assertEqual(
            [cell.value for cell in selection.cells],
            ["value-1", "value-2", "value-3", "value-4"],
        )
        self.assertEqual(selection.cells[0].formula_a1, formula)
        self.assertIsNone(selection.cells[1].formula_a1)
        self.assertEqual(selection.cells[3].formula_a1, "=B2*2")
        self.assertEqual((selection.row_count, selection.column_count), (2, 2))
        self.assertEqual(len(executor.sources), 2)
        self.assertEqual(executor.sources[0], EXCEL_MASK_SCRIPT)
        self.assertIn("get has formula of every cell of selectedRange", executor.sources[0])
        self.assertNotIn("formula2", executor.sources[0].lower())
        self.assertNotIn("special cells", executor.sources[0].lower())

        formula_source = executor.sources[1]
        self.assertIn('{"$B$2", "1", "1"}', formula_source)
        self.assertIn('{"$C$3", "1", "1"}', formula_source)
        self.assertIn("«class 2122» of formulaArea", formula_source)
        self.assertIn("«class F2rc» of formulaArea", formula_source)
        self.assertIn("range formulaAreaAddress of selectedSheet", formula_source)
        self.assertIn('{"$B$2:$C$3", "2", "2"}', formula_source)
        self.assertIn("taggedRangeValues(valueArea", formula_source)
        self.assertIn("get has formula of every cell of targetRange", formula_source)
        self.assertIn(
            "if (count of flatFormulaFlags) is not expectedCellCount then error number -2700",
            formula_source,
        )
        self.assertIn(
            "if isFormulaCell is not true and isFormulaCell is not false then error number -2700",
            formula_source,
        )
        self.assertIn(
            "if isFormulaCell is false and (rawValue is missing value",
            formula_source,
        )
        self.assertIn(
            'evaluate name ("ERROR.TYPE(" & targetAddress & ")")', formula_source
        )
        self.assertNotIn("get text of every cell of targetRange", formula_source)
        self.assertIn('set end of taggedValues to {"blank", ""}', formula_source)
        self.assertIn("SUMPRODUCT(--ISFORMULA(", formula_source)
        self.assertEqual(formula_source.count("my formulaTopologyMatches("), 2)
        self.assertNotIn(workbook, formula_source)
        self.assertNotIn(sheet, formula_source)

    def test_missing_r1c1_is_mapped_to_none(self):
        plan = _parse_mask_result(mask_payload())
        selection = parse_excel_formula_result(
            area_payload([single_area(formula_r1c1=None)]), plan
        )

        self.assertIsNone(selection.cells[0].formula_r1c1)

    def test_blank_and_empty_string_values_remain_distinct(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$B$1",
                rows=1,
                columns=2,
                mask="10",
            )
        )
        selection = parse_excel_formula_result(
            area_payload(
                [single_area('=""', '=""', address="$A$1")],
                address="$A$1:$B$1",
                value_areas=[
                    value_area(["", None], address="$A$1:$B$1", columns=2)
                ],
            ),
            plan,
        )

        self.assertEqual(selection.cells[0].value, "")
        self.assertIsNone(selection.cells[1].value)
        self.assertEqual(selection.cells[0].formula_a1, '=""')
        self.assertIsNone(selection.cells[1].formula_a1)

    def test_reversed_value_chunks_are_reassembled_row_major(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$A$3000",
                rows=3000,
                columns=1,
                mask="1" + ("0" * 2999),
            )
        )
        self.assertEqual(len(plan.value_chunks), 2)
        first_chunk, second_chunk = plan.value_chunks
        value_areas = [
            value_area(
                [f"v{index}" for index in range(2049, 3001)],
                address=second_chunk.address,
                rows=second_chunk.row_count,
                columns=second_chunk.column_count,
            ),
            value_area(
                [f"v{index}" for index in range(1, 2049)],
                address=first_chunk.address,
                rows=first_chunk.row_count,
                columns=first_chunk.column_count,
            ),
        ]

        selection = parse_excel_formula_result(
            area_payload(
                [single_area(address="$A$1")],
                address="$A$1:$A$3000",
                value_areas=value_areas,
            ),
            plan,
        )

        self.assertEqual(len(selection.cells), 3000)
        self.assertEqual(selection.cells[0].value, "v1")
        self.assertEqual(selection.cells[2048].value, "v2049")
        self.assertEqual(selection.cells[-1].value, "v3000")

    def test_compact_bit_mask_is_validated_exactly(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$C$2",
                rows=2,
                columns=3,
                mask="101010",
            )
        )
        self.assertEqual(plan.mask, (True, False, True, False, True, False))

        for invalid_mask in ("10101", "1010100", "10x010"):
            with self.subTest(invalid_mask=invalid_mask):
                with self.assertRaises(ExcelFormulaError) as caught:
                    _parse_mask_result(
                        mask_payload(
                            address="$A$1:$C$2",
                            rows=2,
                            columns=3,
                            mask=invalid_mask,
                        )
                    )
                self.assertEqual(caught.exception.code, INVALID_RESPONSE)


    def test_excel_not_running_short_circuits_before_automation(self):
        executor = FakeExecutor(running=False)

        with self.assertRaises(ExcelFormulaError) as caught:
            read_selected_excel_formulas(executor)

        self.assertEqual(caught.exception.code, EXCEL_NOT_RUNNING)
        self.assertEqual(executor.sources, [])

    def test_content_free_error_codes_remain_distinct(self):
        for code in (
            NO_SELECTION,
            MULTIPLE_AREAS,
            TOO_MANY_CELLS,
            TOO_FRAGMENTED,
            TOO_MUCH_TEXT,
            NO_FORMULAS,
            SELECTION_CHANGED,
            EXECUTION_FAILED,
        ):
            with self.subTest(code=code):
                with self.assertRaises(ExcelFormulaError) as caught:
                    read_selected_excel_formulas(FakeExecutor([["error", code]]))
                self.assertEqual(caught.exception.code, code)
                self.assertEqual(str(caught.exception), code)

    def test_unknown_executor_exception_does_not_leak_message(self):
        executor = FakeExecutor(error=RuntimeError("=PRIVATE_FORMULA(A1)"))

        with self.assertRaises(ExcelFormulaError) as caught:
            read_selected_excel_formulas(executor)

        self.assertEqual(caught.exception.code, EXECUTION_FAILED)
        self.assertNotIn("PRIVATE_FORMULA", str(caught.exception))

    def test_rejects_malformed_mask_payloads(self):
        malformed = (
            None,
            [],
            ["unknown"],
            mask_payload(mask=["false"]),
            mask_payload(address="$A$1:$B$2", rows=2, columns=2, mask=["true"]),
            mask_payload(address="$A$1:$B$2", rows=2, columns=2, mask=["true", "no", "false", "true"]),
            ["error", "not_a_public_code"],
        )
        expected = (
            INVALID_RESPONSE,
            INVALID_RESPONSE,
            INVALID_RESPONSE,
            NO_FORMULAS,
            INVALID_RESPONSE,
            INVALID_RESPONSE,
            EXECUTION_FAILED,
        )
        for payload, code in zip(malformed, expected):
            with self.subTest(payload=payload):
                with self.assertRaises(ExcelFormulaError) as caught:
                    _parse_mask_result(payload)
                self.assertEqual(caught.exception.code, code)

    def test_checkerboard_fifty_formulas_are_all_covered(self):
        flags = [
            "true" if (row + column) % 2 == 0 else "false"
            for row in range(10)
            for column in range(10)
        ]
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$J$10",
                rows=10,
                columns=10,
                mask=flags,
            )
        )
        self.assertEqual(len(plan.chunks), 50)
        areas = []
        for chunk in reversed(plan.chunks):
            formula = f"={chunk.start_row}+{chunk.start_column}"
            areas.append(
                [
                    chunk.address,
                    "1",
                    "1",
                    formula,
                    "present",
                    formula,
                ]
            )

        selection = parse_excel_formula_result(
            area_payload(areas, address="$A$1:$J$10"), plan
        )

        addresses = [cell.address for cell in selection.cells]
        self.assertEqual(len(addresses), 100)
        self.assertEqual(addresses[-3:], ["$H$10", "$I$10", "$J$10"])
        formula_addresses = [
            cell.address for cell in selection.cells if cell.formula_a1 is not None
        ]
        self.assertEqual(len(formula_addresses), 50)
        self.assertEqual(formula_addresses[-3:], ["$F$10", "$H$10", "$J$10"])

    def test_more_than_sixty_four_rectangles_fails_before_formula_read(self):
        flags = [
            "true" if (row + column) % 2 == 0 else "false"
            for row in range(12)
            for column in range(12)
        ]
        executor = FakeExecutor(
            [
                mask_payload(
                    address="$A$1:$L$12",
                    rows=12,
                    columns=12,
                    mask=flags,
                )
            ]
        )

        with self.assertRaises(ExcelFormulaError) as caught:
            read_selected_excel_formulas(executor)

        self.assertEqual(caught.exception.code, TOO_FRAGMENTED)
        self.assertEqual(len(executor.sources), 1)

    def test_dense_selection_is_split_into_bounded_chunks(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$CV$100",
                rows=100,
                columns=100,
                mask=["true"] * 10_000,
            )
        )

        self.assertEqual(len(plan.chunks), 5)
        self.assertTrue(
            all(
                chunk.row_count * chunk.column_count <= MAX_FORMULA_CHUNK_CELLS
                for chunk in (*plan.value_chunks, *plan.chunks)
            )
        )
        self.assertEqual(len(plan.value_chunks), 5)

    def test_vertical_rectangles_are_sorted_globally_by_cell(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$C$2",
                rows=2,
                columns=3,
                mask=["true", "false", "true", "true", "false", "true"],
            )
        )
        # Reverse the area payload to prove output order is coordinate-based.
        areas = [
            ["$C$1:$C$2", "2", "1", [["=3"], ["=6"]], "missing", ""],
            ["$A$1:$A$2", "2", "1", [["=1"], ["=4"]], "missing", ""],
        ]

        selection = parse_excel_formula_result(
            area_payload(areas, address="$A$1:$C$2"), plan
        )

        self.assertEqual(
            [cell.address for cell in selection.cells],
            ["$A$1", "$B$1", "$C$1", "$A$2", "$B$2", "$C$2"],
        )
        self.assertEqual(
            [cell.address for cell in selection.cells if cell.formula_a1],
            ["$A$1", "$C$1", "$A$2", "$C$2"],
        )

    def test_stage_two_error_and_metadata_mismatch_fail_closed(self):
        plan = _parse_mask_result(mask_payload())
        with self.assertRaises(ExcelFormulaError) as changed:
            parse_excel_formula_result(["error", SELECTION_CHANGED], plan)
        self.assertEqual(changed.exception.code, SELECTION_CHANGED)

        with self.assertRaises(ExcelFormulaError) as mismatch:
            parse_excel_formula_result(
                area_payload([single_area()], workbook="Other.xlsx"), plan
            )
        self.assertEqual(mismatch.exception.code, SELECTION_CHANGED)

    def test_stage_two_rejects_extra_missing_or_overlapping_chunks(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$B$1",
                rows=1,
                columns=2,
                mask=["true", "true"],
            )
        )
        invalid_areas = (
            [single_area(address="$A$1")],
            [["$A$1:$B$1", "1", "2", ["=1", "=2"], "missing", ""], single_area(address="$A$1")],
            [["$A$1:$C$1", "1", "3", ["=1", "=2", "=3"], "missing", ""]],
        )
        for areas in invalid_areas:
            with self.subTest(areas=areas):
                with self.assertRaises(ExcelFormulaError) as caught:
                    parse_excel_formula_result(
                        area_payload(areas, address="$A$1:$B$1"), plan
                    )
                self.assertEqual(caught.exception.code, INVALID_RESPONSE)

    def test_stage_two_rejects_missing_overlapping_and_malformed_values(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$B$1",
                rows=1,
                columns=2,
                mask="10",
            )
        )
        invalid_value_areas = (
            [value_area(["one"], address="$A$1")],
            [
                value_area(["one", "two"], address="$A$1:$B$1", columns=2),
                value_area(["duplicate"], address="$A$1"),
            ],
            [["$A$1:$B$1", "1", "2", [["blank", "not-empty"], ["value", "two"]]]],
            [["$A$1:$B$1", "1", "2", [["unknown", "one"], ["value", "two"]]]],
        )
        for value_areas in invalid_value_areas:
            with self.subTest(value_areas=value_areas):
                with self.assertRaises(ExcelFormulaError) as caught:
                    parse_excel_formula_result(
                        area_payload(
                            [single_area(address="$A$1")],
                            address="$A$1:$B$1",
                            value_areas=value_areas,
                        ),
                        plan,
                    )
                self.assertEqual(caught.exception.code, INVALID_RESPONSE)

    def test_formula_text_cap_is_enforced_in_python(self):
        plan = _parse_mask_result(mask_payload())
        with patch("tablemark.excel_formula.MAX_FORMULA_CHARACTERS", 4):
            with self.assertRaises(ExcelFormulaError) as caught:
                parse_excel_formula_result(
                    area_payload([single_area("=1234", None)]), plan
                )
        self.assertEqual(caught.exception.code, TOO_MUCH_TEXT)

    def test_value_text_cap_is_enforced_in_python(self):
        plan = _parse_mask_result(mask_payload())
        with patch("tablemark.excel_formula.MAX_VALUE_CHARACTERS", 3):
            with self.assertRaises(ExcelFormulaError) as caught:
                parse_excel_formula_result(
                    area_payload(
                        [single_area()],
                        value_areas=[value_area(["1234"])],
                    ),
                    plan,
                )
        self.assertEqual(caught.exception.code, TOO_MUCH_TEXT)

    def test_native_executor_maps_automation_denial_by_number_only(self):
        native_script = Mock()
        native_script.executeAndReturnError_.return_value = (
            None,
            {
                "NSAppleScriptErrorNumber": -1743,
                "NSAppleScriptErrorMessage": "=PRIVATE_FORMULA(A1)",
            },
        )
        apple_script_class = Mock()
        apple_script_class.alloc.return_value.initWithSource_.return_value = native_script

        with patch("tablemark.excel_formula.NSAppleScript", apple_script_class):
            with self.assertRaises(ExcelFormulaError) as caught:
                NSAppleScriptExecutor().run(EXCEL_MASK_SCRIPT)

        self.assertEqual(caught.exception.code, AUTOMATION_DENIED)
        self.assertNotIn("PRIVATE_FORMULA", str(caught.exception))


class StableExcelFormulaReaderTests(unittest.TestCase):
    @staticmethod
    def _selection(value: str, formula: str) -> ExcelFormulaSelection:
        return ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (ExcelSelectionCell("$A$1", value, formula, formula),),
        )

    def test_two_equal_reads_return_a_stable_snapshot(self):
        stable = self._selection("2", "=1+1")
        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=[stable, stable],
        ) as reader:
            result = read_stable_selected_excel_formulas()

        self.assertEqual(result, stable)
        self.assertEqual(reader.call_count, 2)

    def test_value_only_recalculation_retries_and_accepts_new_stable_value(self):
        old = self._selection("2", "=1+1")
        new = self._selection("3", "=1+1")
        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=[old, new, new],
        ) as reader:
            result = read_stable_selected_excel_formulas()

        self.assertEqual(result, new)
        self.assertEqual(reader.call_count, 3)

    def test_formula_only_edit_retries_and_accepts_new_stable_formula(self):
        old = self._selection("2", "=1+1")
        new = self._selection("2", "=2*1")
        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=[old, new, new],
        ) as reader:
            result = read_stable_selected_excel_formulas()

        self.assertEqual(result, new)
        self.assertEqual(reader.call_count, 3)

    def test_three_different_snapshots_fail_after_bounded_retry(self):
        snapshots = [
            self._selection("2", "=1+1"),
            self._selection("3", "=1+2"),
            self._selection("4", "=2+2"),
        ]
        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=snapshots,
        ) as reader, self.assertRaises(ExcelFormulaError) as caught:
            read_stable_selected_excel_formulas()

        self.assertEqual(caught.exception.code, SELECTION_CHANGED)
        self.assertEqual(reader.call_count, 3)

    def test_transient_selection_change_retries_but_fatal_error_does_not(self):
        stable = self._selection("2", "=1+1")
        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=[ExcelFormulaError(SELECTION_CHANGED), stable, stable],
        ) as retrying_reader:
            self.assertEqual(read_stable_selected_excel_formulas(), stable)
        self.assertEqual(retrying_reader.call_count, 3)

        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=ExcelFormulaError(AUTOMATION_DENIED),
        ) as fatal_reader, self.assertRaises(ExcelFormulaError) as caught:
            read_stable_selected_excel_formulas()
        self.assertEqual(caught.exception.code, AUTOMATION_DENIED)
        self.assertEqual(fatal_reader.call_count, 1)

    def test_transient_error_resets_the_previous_snapshot(self):
        stable = self._selection("2", "=1+1")
        with patch(
            "tablemark.excel_formula.read_selected_excel_formulas",
            side_effect=[
                stable,
                ExcelFormulaError(SELECTION_CHANGED),
                stable,
            ],
        ) as reader, self.assertRaises(ExcelFormulaError) as caught:
            read_stable_selected_excel_formulas()

        self.assertEqual(caught.exception.code, SELECTION_CHANGED)
        self.assertEqual(reader.call_count, 3)


class FakePasteboard:
    def __init__(self):
        self.declared_types = None
        self.values = {}
        self.clear_count = 0

    def clearContents(self):
        self.clear_count += 1

    def declareTypes_owner_(self, types, owner):
        self.declared_types = list(types)
        self.owner = owner

    def setString_forType_(self, value, pb_type):
        self.values[str(pb_type)] = value


class TextOnlyClipboardTests(unittest.TestCase):
    def test_explicit_writer_drops_every_existing_format(self):
        pasteboard = FakePasteboard()
        pasteboard.values = {
            "com.microsoft.excel.biff12": b"native",
            str(clipboard.NSPasteboardTypeHTML): "<table></table>",
        }
        pasteboard_class = Mock()
        pasteboard_class.generalPasteboard.return_value = pasteboard

        with patch("tablemark.clipboard.NSPasteboard", pasteboard_class):
            clipboard.write_text_only_clipboard("<수식범위 />", mark_generated=True)

        expected_types = {
            str(clipboard.NSPasteboardTypeString),
            clipboard.LEGACY_STRING_TYPE,
            *clipboard.GENERATED_MARKER_TYPES,
        }
        self.assertEqual(pasteboard.clear_count, 1)
        self.assertEqual(set(pasteboard.declared_types), expected_types)
        self.assertNotIn("com.microsoft.excel.biff12", pasteboard.declared_types)
        self.assertNotIn(str(clipboard.NSPasteboardTypeHTML), pasteboard.declared_types)
        for pb_type in (
            str(clipboard.NSPasteboardTypeString),
            clipboard.LEGACY_STRING_TYPE,
        ):
            self.assertEqual(pasteboard.values[pb_type], "<수식범위 />")
        for pb_type, marker in clipboard.GENERATED_MARKER_TYPES.items():
            self.assertEqual(pasteboard.values[pb_type], marker)


class AppFormulaActionTests(unittest.TestCase):
    def test_formula_action_uses_cmd_ctrl_e_globally_and_in_menu(self):
        from AppKit import NSCommandKeyMask, NSControlKeyMask

        from tablemark.app import TabledownApp
        from tablemark.hotkey import CMD, CONTROL, KEY_E, KEY_T, KEY_X

        with (
            patch.object(TabledownApp, "_clear_stale_status_item_visibility"),
            patch("tablemark.app.resolve_language", return_value="en"),
            patch("tablemark.app.load_fill_blanks", return_value=False),
            patch("tablemark.app.load_welcome_shown", return_value=True),
            patch("tablemark.app.login_item.is_supported", return_value=False),
            patch("tablemark.app.clipboard_change_count", return_value=0),
            patch("tablemark.app.threading.Thread"),
            patch("tablemark.app.GlobalHotkeys") as hotkey_class,
        ):
            app = TabledownApp()

        self.assertEqual(app.copy_excel_formulas_item._menuitem.keyEquivalent(), "e")
        self.assertEqual(app.copy_xml_item.title, "Copy selected table as XML")
        self.assertEqual(app.copy_xml_item._menuitem.keyEquivalent(), "x")
        self.assertEqual(
            app.copy_xml_item._menuitem.keyEquivalentModifierMask(),
            NSCommandKeyMask | NSControlKeyMask,
        )
        self.assertEqual(
            app.copy_excel_formulas_item._menuitem.keyEquivalentModifierMask(),
            NSCommandKeyMask | NSControlKeyMask,
        )
        self.assertEqual(
            hotkey_class.return_value.add.call_args_list,
            [
                call(KEY_X, CMD | CONTROL, app.copy_as_xml),
                call(KEY_T, CMD | CONTROL, app.toggle),
                call(
                    KEY_E,
                    CMD | CONTROL,
                    app.copy_selected_excel_formulas,
                ),
            ],
        )
        hotkey_class.return_value.register.assert_called_once_with()

    def test_formula_menu_and_error_translations_exist(self):
        self.assertEqual(
            t("menu.copy_excel_formulas", "ko"), "표의 수식을 포함해 XML로 복사"
        )
        self.assertEqual(
            t("menu.copy_excel_formulas", "en"),
            "Copy table with formulas as XML",
        )
        error_codes = (
            EXCEL_NOT_RUNNING,
            NO_SELECTION,
            MULTIPLE_AREAS,
            TOO_MANY_CELLS,
            TOO_FRAGMENTED,
            TOO_MUCH_TEXT,
            NO_FORMULAS,
            SELECTION_CHANGED,
            AUTOMATION_DENIED,
            EXECUTION_FAILED,
            INVALID_RESPONSE,
        )
        for language in ("ko", "en"):
            self.assertIn(t("menu.copy_excel_formulas", language), t("help.message", language))
            self.assertIn("⌘⌃E", t("help.message", language))
            for code in error_codes:
                key = f"formula.error.{code}"
                self.assertNotEqual(t(key, language), key)

    def test_action_writes_xml_and_flashes_success(self):
        from tablemark.app import TabledownApp

        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (ExcelSelectionCell("$A$1", "2", "=1+1", "=1+1"),),
        )
        fake_app = SimpleNamespace(
            lang="en",
            _flash_icon_success=Mock(),
            _safe_alert=Mock(),
            _clipboard_operation_lock=threading.Lock(),
        )
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log"),
        ):
            TabledownApp.copy_selected_excel_formulas(fake_app, None)

        copied_xml = writer.call_args.args[0]
        self.assertIn('<표범위 통합문서="Book.xlsx"', copied_xml)
        self.assertIn('값="2"', copied_xml)
        self.assertIn('수식="=1+1"', copied_xml)
        writer.assert_called_once_with(copied_xml, mark_generated=True)
        fake_app._flash_icon_success.assert_called_once_with()
        fake_app._safe_alert.assert_not_called()

    def test_action_localizes_content_free_error(self):
        from tablemark.app import TabledownApp

        fake_app = SimpleNamespace(
            lang="ko",
            _flash_icon_success=Mock(),
            _safe_alert=Mock(),
            _clipboard_operation_lock=threading.Lock(),
        )
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_formulas",
                side_effect=ExcelFormulaError(TOO_FRAGMENTED),
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log") as logger,
        ):
            TabledownApp.copy_selected_excel_formulas(fake_app, None)

        writer.assert_not_called()
        logger.assert_called_once_with(
            "copy Excel table with formulas failed: too_fragmented"
        )
        fake_app._safe_alert.assert_called_once_with(
            "Tabledown", t("formula.error.too_fragmented", "ko")
        )


class PackagingTests(unittest.TestCase):
    def test_apple_event_usage_entitlements_and_test_exclusion(self):
        setup_source = (ROOT / "setup.py").read_text(encoding="utf-8")
        self.assertIn('"NSAppleEventsUsageDescription"', setup_source)
        self.assertIn('exclude=("tests", "tests.*")', setup_source)

        with (ROOT / "entitlements" / "mac-app-store.plist").open("rb") as handle:
            store = plistlib.load(handle)
        self.assertTrue(store["com.apple.security.app-sandbox"])
        self.assertTrue(store["com.apple.security.automation.apple-events"])
        self.assertEqual(
            store["com.apple.security.temporary-exception.apple-events"],
            ["com.microsoft.Excel"],
        )
        self.assertEqual(
            store["com.apple.application-identifier"],
            "495S4FVMCB.com.tabledown.app",
        )

        with (ROOT / "entitlements" / "developer-id.plist").open("rb") as handle:
            developer_id = plistlib.load(handle)
        self.assertEqual(
            developer_id,
            {"com.apple.security.automation.apple-events": True},
        )
        build_script = (ROOT / "scripts" / "build_dmg.sh").read_text(
            encoding="utf-8"
        )
        self.assertIn('ENTITLEMENTS="$ROOT/entitlements/developer-id.plist"', build_script)
        self.assertIn('--entitlements "$ENTITLEMENTS" "$APP_STAGE"', build_script)


if __name__ == "__main__":
    unittest.main()
