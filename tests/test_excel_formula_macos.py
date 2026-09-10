from __future__ import annotations

from pathlib import Path
import plistlib
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, call, patch

from Foundation import NSAppleScript

from tablemark import clipboard
from tablemark.converter.formula_export import (
    ExcelFormulaReference,
    ExcelFormulaSelection,
    ExcelReferenceCell,
    ExcelSelectionCell,
    FormulaXmlTooLargeError,
)
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
    REFERENCE_RESULT_STATUS,
    SELECTION_CHANGED,
    TOO_FRAGMENTED,
    TOO_MANY_CELLS,
    TOO_MUCH_TEXT,
    ExcelFormulaError,
    NSAppleScriptExecutor,
    _build_formula_read_script,
    _build_reference_read_script,
    _parse_mask_result,
    _parse_reference_result,
    _rectangle_bounds,
    _reference_requests,
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
    calculation_mode=None,
):
    flags = ["true"] if mask is None else mask
    payload = [
        MASK_RESULT_STATUS,
        workbook,
        sheet,
        address,
        str(rows),
        str(columns),
        flags,
    ]
    if calculation_mode is not None:
        payload.append(calculation_mode)
    return payload


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


def reference_payload(
    references,
    *,
    workbook="Book.xlsx",
    sheet="Sheet1",
    address="$A$1",
):
    return [
        REFERENCE_RESULT_STATUS,
        workbook,
        sheet,
        address,
        references,
    ]


def reference_area(
    owner,
    sheet,
    address,
    values,
    *,
    rows=1,
    columns=1,
):
    return [
        "ok",
        owner,
        sheet,
        address,
        str(rows),
        str(columns),
        [
            ["blank", ""] if value is None else ["value", value]
            for value in values
        ],
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


class NativeExcelRunningTests(unittest.TestCase):
    @staticmethod
    def _application(bundle_id="com.microsoft.Excel", *, terminated=False):
        app = Mock()
        app.bundleIdentifier.return_value = bundle_id
        app.isTerminated.return_value = terminated
        return app

    def test_primary_match_keeps_fast_path_without_workspace_or_automation(self):
        with (
            patch("tablemark.excel_formula.NSRunningApplication") as running,
            patch("tablemark.excel_formula.NSWorkspace") as workspace,
            patch("tablemark.excel_formula.NSAppleScript") as apple_script,
        ):
            running.runningApplicationsWithBundleIdentifier_.return_value = [object()]

            self.assertTrue(NSAppleScriptExecutor().is_excel_running())

        running.runningApplicationsWithBundleIdentifier_.assert_called_once_with(
            "com.microsoft.Excel"
        )
        workspace.sharedWorkspace.assert_not_called()
        apple_script.alloc.assert_not_called()

    def test_empty_primary_uses_live_exact_bundle_from_workspace(self):
        live_excel = self._application()
        with (
            patch("tablemark.excel_formula.NSRunningApplication") as running,
            patch("tablemark.excel_formula.NSWorkspace") as workspace,
            patch("tablemark.excel_formula.NSAppleScript") as apple_script,
        ):
            running.runningApplicationsWithBundleIdentifier_.return_value = []
            workspace.sharedWorkspace.return_value.runningApplications.return_value = [
                self._application("com.microsoft.Excel.helper"),
                self._application(terminated=True),
                live_excel,
            ]

            self.assertTrue(NSAppleScriptExecutor().is_excel_running())

        workspace.sharedWorkspace.return_value.runningApplications.assert_called_once_with()
        live_excel.isTerminated.assert_called_once_with()
        apple_script.alloc.assert_not_called()

    def test_workspace_missing_unrelated_and_terminated_apps_remain_unavailable(self):
        for applications in (
            [],
            [self._application("com.microsoft.Excel.helper")],
            [self._application(terminated=True)],
        ):
            with (
                self.subTest(applications=applications),
                patch("tablemark.excel_formula.NSRunningApplication") as running,
                patch("tablemark.excel_formula.NSWorkspace") as workspace,
                patch("tablemark.excel_formula.NSAppleScript") as apple_script,
            ):
                running.runningApplicationsWithBundleIdentifier_.return_value = []
                workspace.sharedWorkspace.return_value.runningApplications.return_value = applications

                self.assertFalse(NSAppleScriptExecutor().is_excel_running())

                apple_script.alloc.assert_not_called()

    def test_workspace_failure_is_not_reported_as_a_running_application(self):
        with (
            patch("tablemark.excel_formula.NSRunningApplication") as running,
            patch("tablemark.excel_formula.NSWorkspace") as workspace,
            patch("tablemark.excel_formula.NSAppleScript") as apple_script,
        ):
            running.runningApplicationsWithBundleIdentifier_.return_value = []
            workspace.sharedWorkspace.return_value.runningApplications.side_effect = RuntimeError()

            with self.assertRaises(RuntimeError):
                NSAppleScriptExecutor().is_excel_running()

        apple_script.alloc.assert_not_called()

    def test_native_unavailable_reader_never_invokes_excel_automation(self):
        with (
            patch("tablemark.excel_formula.NSRunningApplication") as running,
            patch("tablemark.excel_formula.NSWorkspace") as workspace,
            patch("tablemark.excel_formula.NSAppleScript") as apple_script,
        ):
            running.runningApplicationsWithBundleIdentifier_.return_value = []
            workspace.sharedWorkspace.return_value.runningApplications.return_value = []

            with self.assertRaises(ExcelFormulaError) as caught:
                read_selected_excel_formulas(NSAppleScriptExecutor())

        self.assertEqual(caught.exception.code, EXCEL_NOT_RUNNING)
        apple_script.alloc.assert_not_called()


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

        reference_selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (ExcelSelectionCell("$A$1", "2", "=Sheet2!B1+A2"),),
        )
        requests, _completeness = _reference_requests(reference_selection)

        for source in (
            EXCEL_MASK_SCRIPT,
            _build_formula_read_script(plan),
            _build_reference_read_script(
                _parse_mask_result(mask_payload()), requests
            ),
        ):
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
            calculation_mode="calculation manual",
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
        third = reference_payload(
            [
                reference_area("$B$2", sheet, "$A$1", ["source-a1"]),
                reference_area("$C$3", sheet, "$B$2", ["value-1"]),
            ],
            workbook=workbook,
            sheet=sheet,
            address="$B$2:$C$3",
        )
        executor = FakeExecutor([first, second, third])

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
        self.assertEqual(selection.cells[0].references[0].cells[0].value, "source-a1")
        self.assertFalse(selection.cells[0].references_complete)
        self.assertIsNone(selection.cells[1].formula_a1)
        self.assertEqual(selection.cells[3].formula_a1, "=B2*2")
        self.assertEqual(selection.cells[3].references[0].cells[0].value, "value-1")
        self.assertEqual((selection.row_count, selection.column_count), (2, 2))
        self.assertEqual(selection.calculation_mode, "manual")
        self.assertEqual(selection.calculation_state, "unavailable")
        self.assertEqual(len(executor.sources), 3)
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
        self.assertIn("if expectedCellCount is 1 then", formula_source)
        self.assertIn("set rawValue to get value2 of targetRange", formula_source)
        self.assertIn(
            "set rawValues to get value2 of every cell of targetRange",
            formula_source,
        )
        self.assertNotIn("get value of targetRange", formula_source)
        self.assertIn("set rawFormulaFlag to get has formula of targetRange", formula_source)
        self.assertIn('set end of taggedValues to {"number", rawValue}', formula_source)
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
        reference_source = executor.sources[2]
        self.assertIn('"$B$2"', reference_source)
        self.assertIn('"$A$1"', reference_source)
        self.assertIn("taggedReferenceValues", reference_source)
        self.assertIn("if expectedCellCount is 1 then", reference_source)
        self.assertIn("set rawValue to get value2 of targetRange", reference_source)
        self.assertIn(
            "set rawValues to get value2 of every cell of targetRange",
            reference_source,
        )
        self.assertNotIn("get value of targetRange", reference_source)
        self.assertIn('set end of taggedValues to {"number", rawValue}', reference_source)
        self.assertIn("on excelReferenceErrorText", reference_source)
        self.assertIn(
            'evaluate name ("ERROR.TYPE(" & targetAddress & ")")',
            reference_source,
        )
        self.assertIn("reference style A1 external true", reference_source)
        self.assertNotIn(workbook, reference_source)
        self.assertNotIn(sheet, reference_source)

    def test_same_and_cross_sheet_reference_values_are_attached(self):
        first = mask_payload(
            address="$E$6",
            mask="1",
            workbook="Budget.xlsx",
            sheet="기본_표",
        )
        second = area_payload(
            [single_area("=C6*'단가 표'!D6", "=RC[-2]*'단가 표'!R6C4", address="$E$6")],
            address="$E$6",
            workbook="Budget.xlsx",
            sheet="기본_표",
            value_areas=[value_area(["120"], address="$E$6")],
        )
        third = reference_payload(
            [
                reference_area("$E$6", "기본_표", "$C$6", ["10"]),
                reference_area("$E$6", "단가 표", "$D$6", ["12"]),
            ],
            address="$E$6",
            workbook="Budget.xlsx",
            sheet="기본_표",
        )

        selection = read_selected_excel_formulas(FakeExecutor([first, second, third]))

        formula_cell = selection.cells[0]
        self.assertEqual(formula_cell.value, "120")
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

    def test_unresolved_dynamic_reference_keeps_formula_and_marks_partial(self):
        first = mask_payload(address="$A$1", mask="1")
        second = area_payload(
            [single_area('=INDIRECT("Sheet2!A1")')],
            value_areas=[value_area(["10"])],
        )
        executor = FakeExecutor([first, second])

        selection = read_selected_excel_formulas(executor)

        self.assertEqual(selection.cells[0].formula_a1, '=INDIRECT("Sheet2!A1")')
        self.assertEqual(selection.cells[0].references, ())
        self.assertFalse(selection.cells[0].references_complete)
        self.assertEqual(
            selection.cells[0].reference_issues,
            ("dynamic_reference",),
        )
        self.assertEqual(len(executor.sources), 2)

    def test_unavailable_static_reference_keeps_formula_and_marks_partial(self):
        first = mask_payload(address="$A$1", mask="1")
        second = area_payload(
            [single_area("=Sheet2!A1")],
            value_areas=[value_area(["10"])],
        )
        third = reference_payload(
            [["unresolved", "$A$1", "Sheet2", "$A$1"]]
        )

        selection = read_selected_excel_formulas(FakeExecutor([first, second, third]))

        self.assertEqual(selection.cells[0].formula_a1, "=Sheet2!A1")
        self.assertEqual(selection.cells[0].references, ())
        self.assertFalse(selection.cells[0].references_complete)
        self.assertEqual(
            selection.cells[0].reference_issues,
            ("read_failed",),
        )

    def test_reference_limits_keep_bounded_prefix_and_mark_partial(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$C$1",
            1,
            1,
            (ExcelSelectionCell("$C$1", "3", "=A1+B1"),),
        )

        with patch("tablemark.excel_formula.MAX_REFERENCE_RANGES", 1):
            requests, completeness = _reference_requests(selection)

        self.assertEqual(
            [(request.target.sheet, request.target.address) for request in requests],
            [("Sheet1", "$A$1")],
        )
        self.assertFalse(completeness["$C$1"])

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

    def test_numeric_tags_expand_exponents_without_changing_literal_text(self):
        plan = _parse_mask_result(
            mask_payload(
                address="$A$1:$C$1",
                rows=1,
                columns=3,
                mask="100",
            )
        )
        selection = parse_excel_formula_result(
            area_payload(
                [single_area("=1", "=1", address="$A$1")],
                address="$A$1:$C$1",
                value_areas=[
                    [
                        "$A$1:$C$1",
                        "1",
                        "3",
                        [
                            ["number", " 3.15E+5"],
                            ["number", "1.2e-7"],
                            ["text", "3.15E+5"],
                        ],
                    ]
                ],
            ),
            plan,
        )

        self.assertEqual(
            [cell.value for cell in selection.cells],
            ["315000", "0.00000012", "3.15E+5"],
        )
        self.assertEqual(
            [cell.value_kind for cell in selection.cells],
            ["number", "number", "text"],
        )

    def test_shared_reference_is_planned_once_for_every_owner(self):
        cells = tuple(
            ExcelSelectionCell(
                f"$B${row}",
                "1",
                "=$A$1",
            )
            for row in range(1, 301)
        )
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1:$B$300",
            300,
            1,
            cells,
        )

        requests, completeness = _reference_requests(selection)

        self.assertEqual(len(requests), 1)
        self.assertEqual(requests[0].target.address, "$A$1")
        self.assertEqual(len(requests[0].owner_addresses), 300)
        self.assertTrue(all(completeness.values()))

    def test_shared_reference_read_preserves_each_owner_formula_order(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1:$B$2",
            2,
            1,
            (
                ExcelSelectionCell("$B$1", "4", "=A1+C1"),
                ExcelSelectionCell("$B$2", "4", "=C1+A1"),
            ),
        )
        requests, completeness = _reference_requests(selection)
        payload = reference_payload(
            [
                reference_area(
                    request.owner_address,
                    request.target.sheet,
                    request.target.address,
                    ["2"],
                )
                for request in requests
            ],
            address="$B$1:$B$2",
        )

        enriched = _parse_reference_result(
            payload,
            selection,
            requests,
            completeness,
            {},
        )

        self.assertEqual(
            [reference.address for reference in enriched.cells[0].references],
            ["$A$1", "$C$1"],
        )
        self.assertEqual(
            [reference.address for reference in enriched.cells[1].references],
            ["$C$1", "$A$1"],
        )

    def test_reference_value_limit_is_reported_distinctly(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1",
            1,
            1,
            (ExcelSelectionCell("$B$1", "1", "=A1"),),
        )
        requests, completeness = _reference_requests(selection)
        request = requests[0]

        enriched = _parse_reference_result(
            reference_payload(
                [
                    [
                        "too_large",
                        request.owner_address,
                        request.target.sheet,
                        request.target.address,
                    ]
                ],
                address="$B$1",
            ),
            selection,
            requests,
            completeness,
            {},
        )

        self.assertFalse(enriched.cells[0].references_complete)
        self.assertEqual(
            enriched.cells[0].reference_issues,
            ("value_size_limit",),
        )

    def test_reference_value_types_are_kept_for_safe_substitution(self):
        first = mask_payload(address="$B$1", mask="1")
        second = area_payload(
            [single_area("=A1", address="$B$1")],
            address="$B$1",
            value_areas=[value_area(["315000"], address="$B$1")],
        )
        third = reference_payload(
            [
                [
                    "ok",
                    "$B$1",
                    "Sheet1",
                    "$A$1",
                    "1",
                    "1",
                    [["number", "3.15E+5"]],
                ]
            ],
            address="$B$1",
        )

        selection = read_selected_excel_formulas(FakeExecutor([first, second, third]))
        reference_cell = selection.cells[0].references[0].cells[0]

        self.assertEqual(reference_cell.value, "315000")
        self.assertEqual(reference_cell.value_kind, "number")

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
    def _selection(
        value: str, formula: str, *, reference_value: str | None = None
    ) -> ExcelFormulaSelection:
        references = ()
        if reference_value is not None:
            references = (
                ExcelFormulaReference(
                    "Sheet2",
                    "$B$2",
                    (ExcelReferenceCell("$B$2", reference_value),),
                ),
            )
        return ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$A$1", value, formula, formula, references=references
                ),
            ),
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

    def test_reference_value_change_retries_and_accepts_new_stable_snapshot(self):
        old = self._selection("2", "=Sheet2!B2", reference_value="1")
        new = self._selection("2", "=Sheet2!B2", reference_value="2")
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
        self.change_count = 0

    def clearContents(self):
        self.clear_count += 1
        self.change_count += 1
        self.values = {}

    def declareTypes_owner_(self, types, owner):
        self.declared_types = list(types)
        self.owner = owner

    def setString_forType_(self, value, pb_type):
        self.values[str(pb_type)] = value
        return True

    def stringForType_(self, pb_type):
        return self.values.get(str(pb_type))

    def changeCount(self):
        return self.change_count

    def types(self):
        return list(self.values)

    def dataForType_(self, pb_type):
        return self.values.get(str(pb_type))


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

    def test_expected_change_count_mismatch_does_not_clear_clipboard(self):
        pasteboard = FakePasteboard()
        pasteboard.change_count = 9
        pasteboard.values = {str(clipboard.NSPasteboardTypeString): "new"}
        pasteboard_class = Mock()
        pasteboard_class.generalPasteboard.return_value = pasteboard

        with (
            patch("tablemark.clipboard.NSPasteboard", pasteboard_class),
            self.assertRaises(clipboard.ClipboardChangedError),
        ):
            clipboard.write_text_only_clipboard(
                "<표 />",
                expected_change_count=8,
            )

        self.assertEqual(pasteboard.clear_count, 0)
        self.assertEqual(
            pasteboard.values[str(clipboard.NSPasteboardTypeString)],
            "new",
        )

    def test_preserving_writer_rechecks_generation_before_clearing(self):
        pasteboard = FakePasteboard()
        pasteboard.change_count = 9
        pasteboard.values = {
            str(clipboard.NSPasteboardTypeString): "new",
            "com.microsoft.excel.biff12": b"native",
        }
        original_data_for_type = pasteboard.dataForType_

        def change_during_preservation(pb_type):
            data = original_data_for_type(pb_type)
            pasteboard.change_count = 10
            return data

        pasteboard.dataForType_ = change_during_preservation
        pasteboard_class = Mock()
        pasteboard_class.generalPasteboard.return_value = pasteboard

        with (
            patch("tablemark.clipboard.NSPasteboard", pasteboard_class),
            self.assertRaises(clipboard.ClipboardChangedError),
        ):
            clipboard.write_clipboard(
                text="old conversion",
                expected_change_count=9,
            )

        self.assertEqual(pasteboard.clear_count, 0)
        self.assertEqual(
            pasteboard.values[str(clipboard.NSPasteboardTypeString)],
            "new",
        )
        self.assertEqual(
            pasteboard.values["com.microsoft.excel.biff12"],
            b"native",
        )

    def test_failed_write_or_readback_is_not_reported_as_success(self):
        pasteboard = FakePasteboard()
        pasteboard_class = Mock()
        pasteboard_class.generalPasteboard.return_value = pasteboard

        def fail_marker(_value, pb_type):
            if str(pb_type) == clipboard.TABLEDOWN_GENERATED_TYPE:
                return False
            pasteboard.values[str(pb_type)] = _value
            return True

        pasteboard.setString_forType_ = fail_marker
        with (
            patch("tablemark.clipboard.NSPasteboard", pasteboard_class),
            self.assertRaises(clipboard.ClipboardWriteError),
        ):
            clipboard.write_text_only_clipboard("<표 />")

    def test_readback_mismatch_is_not_reported_as_success(self):
        pasteboard = FakePasteboard()
        pasteboard_class = Mock()
        pasteboard_class.generalPasteboard.return_value = pasteboard
        original_read = pasteboard.stringForType_

        def wrong_legacy_text(pb_type):
            if str(pb_type) == clipboard.LEGACY_STRING_TYPE:
                return "<다른표 />"
            return original_read(pb_type)

        pasteboard.stringForType_ = wrong_legacy_text
        with (
            patch("tablemark.clipboard.NSPasteboard", pasteboard_class),
            self.assertRaises(clipboard.ClipboardWriteError),
        ):
            clipboard.write_text_only_clipboard("<표 />")


class AppFormulaActionTests(unittest.TestCase):
    @staticmethod
    def _app(lang="en"):
        from tablemark.app import TabledownApp

        app = SimpleNamespace(
            fill_blanks=False,
            lang=lang,
            _flash_icon_success=Mock(),
            _safe_alert=Mock(),
            _clipboard_operation_lock=threading.Lock(),
            _explicit_export_lock=threading.Lock(),
            _explicit_export_active=None,
            _stop_watcher=threading.Event(),
            copy_xml_item=SimpleNamespace(title="", _menuitem=Mock()),
            copy_excel_formulas_item=SimpleNamespace(title="", _menuitem=Mock()),
            copy_markdown_item=SimpleNamespace(title="", _menuitem=Mock()),
        )
        for name in (
            "copy_as_markdown",
            "copy_as_xml",
            "copy_selected_excel_formulas",
            "_start_explicit_export",
            "_run_explicit_export",
            "_perform_explicit_export",
            "_finish_explicit_export",
            "_alert_explicit_export_error",
            "_set_explicit_export_busy",
            "_update_explicit_export_menu",
        ):
            setattr(app, name, getattr(TabledownApp, name).__get__(app))
        return app

    @staticmethod
    def _run(app):
        with (
            patch("tablemark.app.clipboard_change_count", return_value=23),
            patch(
                "tablemark.app.AppHelper.callAfter",
                side_effect=lambda callback, *args: callback(*args),
            ),
        ):
            app.copy_selected_excel_formulas(None)

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
        self.assertEqual(
            app.copy_xml_item.title,
            "Copy as XML",
        )
        self.assertEqual(app.copy_xml_item._menuitem.keyEquivalent(), "x")
        self.assertEqual(app.copy_markdown_item.title, "Copy Markdown")
        self.assertEqual(app.copy_excel_formulas_item.title, "Copy as XML with Formulas")
        self.assertEqual(app.copy_markdown_item._menuitem.keyEquivalent(), "")
        menu_titles = [getattr(item, "title", None) for item in app.menu.values()]
        copy_titles = [title for title in menu_titles if title and title.startswith("Copy")]
        self.assertEqual(
            copy_titles,
            ["Copy Markdown", "Copy as XML", "Copy as XML with Formulas"],
        )
        self.assertNotIn("Copy compact XML for AI", menu_titles)
        self.assertFalse(hasattr(app, "copy_ai_formulas_item"))
        self.assertFalse(hasattr(app, "copy_selected_excel_formulas_for_ai"))
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
        menu_labels = {
            "ko": ("마크다운 복사", "XML 변환 복사", "수식 포함 XML 변환 복사"),
            "en": ("Copy Markdown", "Copy as XML", "Copy as XML with Formulas"),
        }
        for language, expected in menu_labels.items():
            self.assertEqual(
                tuple(t(key, language) for key in (
                    "menu.copy_markdown", "menu.copy_xml", "menu.copy_excel_formulas",
                )),
                expected,
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
            "clipboard_changed",
            "clipboard_write_failed",
            "output_too_large",
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
        fake_app = self._app("en")
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log"),
        ):
            self._run(fake_app)

        copied_xml = writer.call_args.args[0]
        self.assertIn('<표범위 통합문서="Book.xlsx"', copied_xml)
        self.assertIn('값="2"', copied_xml)
        self.assertIn('수식="=1+1"', copied_xml)
        writer.assert_called_once_with(
            copied_xml,
            mark_generated=True,
            expected_change_count=23,
        )
        fake_app._flash_icon_success.assert_called_once_with()
        fake_app._safe_alert.assert_not_called()

    @staticmethod
    def _notice_selection(*, partial=True, calculation_state="done"):
        formula = '=INDIRECT("A1")' if partial else "=1+1"
        return ExcelFormulaSelection(
            "PrivateBook.xlsx", "Sheet1", "$B$1", 1, 1,
            (ExcelSelectionCell(
                "$B$1", "2", formula_a1=formula,
                references_complete=not partial, value_kind="number",
                reference_issues=("dynamic_reference",) if partial else (),
            ),),
            calculation_mode="automatic", calculation_state=calculation_state,
        )

    def test_copy_notice_follows_verified_write_and_preserves_formula(self):
        cases = (
            (True, "done", "partial_references"),
            (False, "pending", "calculation_incomplete"),
            (True, "calculating", "partial_references_and_calculation"),
        )
        for language in ("ko", "en"):
            for partial, state, notice in cases:
                with self.subTest(language=language, notice=notice):
                    selection = self._notice_selection(
                        partial=partial, calculation_state=state,
                    )
                    app = self._app(language)
                    events = []
                    app._safe_alert.side_effect = lambda *_: events.append("notice")
                    with (
                        patch("tablemark.app.read_stable_selected_excel_formulas", return_value=selection),
                        patch("tablemark.app.write_text_only_clipboard", side_effect=lambda *_, **__: events.append("write")) as writer,
                        patch("tablemark.app.log") as logger,
                    ):
                        self._run(app)

                    copied_xml = writer.call_args.args[0]
                    self.assertIn('값="2"', copied_xml)
                    self.assertIn('주소="$B$1"', copied_xml)
                    if partial:
                        self.assertIn('참조상태="일부"', copied_xml)
                        self.assertIn('수식="=INDIRECT(&quot;A1&quot;)"', copied_xml)
                    else:
                        self.assertIn('수식="=1+1"', copied_xml)
                    if state != "done":
                        self.assertIn('계산결과상태="calculation_incomplete"', copied_xml)
                    self.assertEqual(events, ["write", "notice"])
                    key = f"formula.copy_notice.{notice}"
                    self.assertNotEqual(t(key, language), key)
                    app._safe_alert.assert_called_once_with(
                        t("formula.copy_notice_title", language), t(key, language),
                    )
                    app._flash_icon_success.assert_called_once_with()
                    self.assertNotIn("PrivateBook", str(app._safe_alert.call_args))
                    self.assertNotIn("PrivateBook", str(logger.call_args_list))
                    self.assertFalse(app._explicit_export_lock.locked())
                    self.assertIsNone(app._explicit_export_active)

    def test_copy_notice_does_not_override_write_failure_or_new_copy(self):
        cases = (
            (clipboard.ClipboardWriteError, "clipboard_write_failed"),
            (clipboard.ClipboardChangedError, "clipboard_changed"),
        )
        for error_class, code in cases:
            with self.subTest(code=code):
                app = self._app("ko")
                with (
                    patch("tablemark.app.read_stable_selected_excel_formulas", return_value=self._notice_selection()),
                    patch("tablemark.app.write_text_only_clipboard", side_effect=error_class("content-free")),
                    patch("tablemark.app.log"),
                ):
                    self._run(app)
                app._safe_alert.assert_called_once_with(
                    "Tabledown", t(f"formula.error.{code}", "ko"),
                )
                app._flash_icon_success.assert_not_called()
                self.assertFalse(app._explicit_export_lock.locked())
                self.assertIsNone(app._explicit_export_active)

    def test_copy_notice_is_suppressed_when_quitting(self):
        app = self._app("ko")
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=self._notice_selection()),
            patch("tablemark.app.write_text_only_clipboard", side_effect=lambda *_, **__: app._stop_watcher.set()),
            patch("tablemark.app.log"),
        ):
            self._run(app)
        app._safe_alert.assert_not_called()
        app._flash_icon_success.assert_not_called()
        self.assertFalse(app._explicit_export_lock.locked())

    def test_partial_copy_can_be_retried_after_write_failure(self):
        app = self._app("ko")
        with (
            patch("tablemark.app.read_stable_selected_excel_formulas", return_value=self._notice_selection()),
            patch("tablemark.app.write_text_only_clipboard", side_effect=[clipboard.ClipboardWriteError("content-free"), None]) as writer,
            patch("tablemark.app.log"),
        ):
            self._run(app)
            app._flash_icon_success.assert_not_called()
            self.assertFalse(app._explicit_export_lock.locked())
            self._run(app)
        self.assertEqual(writer.call_count, 2)
        app._flash_icon_success.assert_called_once_with()
        self.assertEqual(app._safe_alert.call_args_list, [
            call("Tabledown", t("formula.error.clipboard_write_failed", "ko")),
            call(t("formula.copy_notice_title", "ko"), t("formula.copy_notice.partial_references", "ko")),
        ])
        self.assertFalse(app._explicit_export_lock.locked())

    def test_action_localizes_content_free_error(self):
        from tablemark.app import TabledownApp

        fake_app = self._app("ko")
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_formulas",
                side_effect=ExcelFormulaError(TOO_FRAGMENTED),
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(fake_app)

        writer.assert_not_called()
        logger.assert_called_once_with(
            "copy Excel table with formulas failed: too_fragmented"
        )
        fake_app._safe_alert.assert_called_once_with(
            "Tabledown", t("formula.error.too_fragmented", "ko")
        )

    def test_action_reports_output_limit_before_writing_clipboard(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (ExcelSelectionCell("$A$1", "2", "=1+1", "=1+1"),),
        )
        fake_app = self._app("ko")
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_formulas",
                return_value=selection,
            ),
            patch(
                "tablemark.app.formula_selection_to_ai_xml",
                side_effect=FormulaXmlTooLargeError("content-free"),
            ),
            patch("tablemark.app.write_text_only_clipboard") as writer,
            patch("tablemark.app.log") as logger,
        ):
            self._run(fake_app)

        writer.assert_not_called()
        logger.assert_called_once_with(
            "copy Excel table with formulas failed: output_too_large"
        )
        fake_app._safe_alert.assert_called_once_with(
            "Tabledown", t("formula.error.output_too_large", "ko")
        )
        fake_app._flash_icon_success.assert_not_called()

    def test_action_reports_clipboard_write_failure_without_success(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (ExcelSelectionCell("$A$1", "2", "=1+1", "=1+1"),),
        )
        fake_app = self._app("ko")
        with (
            patch(
                "tablemark.app.read_stable_selected_excel_formulas",
                return_value=selection,
            ),
            patch(
                "tablemark.app.write_text_only_clipboard",
                side_effect=clipboard.ClipboardWriteError("content-free"),
            ),
            patch("tablemark.app.log") as logger,
        ):
            self._run(fake_app)

        logger.assert_called_once_with(
            "copy Excel table with formulas failed: clipboard_write_failed"
        )
        fake_app._safe_alert.assert_called_once_with(
            "Tabledown", t("formula.error.clipboard_write_failed", "ko")
        )
        fake_app._flash_icon_success.assert_not_called()


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
