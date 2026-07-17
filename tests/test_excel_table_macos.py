"""macOS direct Excel-selection table reader regression tests."""

from __future__ import annotations

from pathlib import Path
import unittest
import xml.etree.ElementTree as ET

from Foundation import NSAppleScript

from tablemark.converter.html_to_md import html_table_to_model
from tablemark.converter.table_xml import model_to_xml
from tablemark.excel_formula import (
    _AE_LIST_DESCRIPTOR_TYPE,
    _descriptor_to_python,
    DISPLAY_OVERFLOW,
    EXCEL_NOT_RUNNING,
    EXECUTION_FAILED,
    INVALID_RESPONSE,
    PARTIAL_MERGE,
    SELECTION_CHANGED,
    TOO_MUCH_TEXT,
    ExcelFormulaError,
)
from tablemark.excel_table import (
    EXCEL_TABLE_SCRIPT,
    MAX_HASH_LITERAL_FALLBACK_CELLS,
    MAX_TABLE_SELECTION_CELLS,
    TABLE_RESULT_STATUS,
    _VALUE_DELIMITER,
    _VALUE_ESCAPE,
    ExcelTableSelection,
    excel_table_selection_to_html,
    excel_table_selection_to_model,
    parse_excel_table_result,
    read_selected_excel_table,
    read_stable_selected_excel_table,
)


def table_payload(
    values,
    *,
    workbook="Book.xlsx",
    sheet="Sheet1",
    address="$A$1:$B$2",
    rows=2,
    columns=2,
    merges=(),
):
    value_texts = ["" if value is None else value for value in values]
    encoded_values = [
        value.replace(_VALUE_ESCAPE, _VALUE_ESCAPE + "e").replace(
            _VALUE_DELIMITER, _VALUE_ESCAPE + "d"
        )
        for value in value_texts
    ]
    return [
        TABLE_RESULT_STATUS,
        workbook,
        sheet,
        address,
        str(rows),
        str(columns),
        "".join("b" if value is None else "v" for value in values),
        _VALUE_DELIMITER.join(encoded_values),
        list(merges),
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


class ExcelTableReaderTests(unittest.TestCase):
    def test_native_empty_apple_event_list_decodes_as_empty_list(self):
        class EmptyListDescriptor:
            @staticmethod
            def numberOfItems():
                return 0

            @staticmethod
            def descriptorType():
                return _AE_LIST_DESCRIPTOR_TYPE

            @staticmethod
            def stringValue():
                return None

        self.assertEqual(_descriptor_to_python(EmptyListDescriptor()), [])

    @unittest.skipUnless(
        Path("/Applications/Microsoft Excel.app").exists(),
        "Microsoft Excel terminology is required to compile the script",
    )
    def test_generated_applescript_compiles(self):
        script = NSAppleScript.alloc().initWithSource_(EXCEL_TABLE_SCRIPT)
        compiled, error_info = script.compileAndReturnError_(None)
        self.assertTrue(compiled, error_info)

    def test_reader_accepts_constant_only_table_and_reads_merges(self):
        payload = table_payload(
            ["구분", "값", "A", "10"],
            merges=["$A$1:$B$1"],
        )
        executor = FakeExecutor([payload])

        selection = read_selected_excel_table(executor)

        self.assertEqual(selection.values, ("구분", "값", "A", "10"))
        self.assertEqual(selection.merge_areas, ("$A$1:$B$1",))
        self.assertEqual(executor.sources, [EXCEL_TABLE_SCRIPT])
        self.assertNotIn("has formula", EXCEL_TABLE_SCRIPT.lower())
        self.assertIn(
            "get string value of targetRange", EXCEL_TABLE_SCRIPT
        )
        self.assertIn("set rawValues to get value of targetRange", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("value of every cell of targetRange", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("get number format of every cell", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("evaluate name formatExpression", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("get text of every cell of targetRange", EXCEL_TABLE_SCRIPT)
        self.assertIn(
            "get merge area of every cell of selectedRange", EXCEL_TABLE_SCRIPT
        )
        self.assertIn(
            "get address of every item of rawMergeAreas", EXCEL_TABLE_SCRIPT
        )
        self.assertNotIn("get merge cells of every cell", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("is not in mergedAddresses", EXCEL_TABLE_SCRIPT)
        self.assertIn('set valueKinds to ""', EXCEL_TABLE_SCRIPT)
        self.assertNotIn('{{"value",', EXCEL_TABLE_SCRIPT)
        self.assertIn("joinValueTexts(valueTexts)", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("isLiteralHashText", EXCEL_TABLE_SCRIPT)
        self.assertIn(
            "if (count of hashCandidateIndexes) > 32 then",
            EXCEL_TABLE_SCRIPT,
        )
        self.assertEqual(MAX_HASH_LITERAL_FALLBACK_CELLS, 32)
        self.assertIn(f'{{"error", "{DISPLAY_OVERFLOW}"}}', EXCEL_TABLE_SCRIPT)
        self.assertEqual(MAX_TABLE_SELECTION_CELLS, 10_000)
        self.assertIn("selectedCellCount > 10000", EXCEL_TABLE_SCRIPT)
        self.assertNotIn("clipboard", EXCEL_TABLE_SCRIPT.lower())

    def test_blank_and_empty_string_remain_distinct(self):
        selection = parse_excel_table_result(
            table_payload(["머리", "값", None, ""], rows=2, columns=2)
        )

        self.assertIsNone(selection.values[2])
        self.assertEqual(selection.values[3], "")

    def test_compact_value_blob_roundtrips_reserved_characters(self):
        selection = parse_excel_table_result(
            table_payload(
                [
                    "머리",
                    "값",
                    f"앞{_VALUE_DELIMITER}뒤",
                    f"앞{_VALUE_ESCAPE}뒤",
                ]
            )
        )

        self.assertEqual(
            selection.values[2:],
            (f"앞{_VALUE_DELIMITER}뒤", f"앞{_VALUE_ESCAPE}뒤"),
        )

    def test_merge_order_is_canonical_for_snapshot_comparison(self):
        selection = parse_excel_table_result(
            table_payload(
                ["A", None, "B", None],
                merges=["$B$1:$B$2", "$A$1:$A$2"],
            )
        )

        self.assertEqual(
            selection.merge_areas,
            ("$A$1:$A$2", "$B$1:$B$2"),
        )

    def test_batch_merge_addresses_filter_singletons_and_duplicates(self):
        selection = parse_excel_table_result(
            table_payload(
                ["A", "B", "C", None],
                merges=["$A$1", "$B$1", "$A$2:$B$2", "$A$2:$B$2"],
            )
        )

        self.assertEqual(selection.merge_areas, ("$A$2:$B$2",))

    def test_partial_merge_fails_closed(self):
        with self.assertRaisesRegex(ExcelFormulaError, PARTIAL_MERGE):
            parse_excel_table_result(
                table_payload(
                    ["머리", "값", "A", "1"],
                    merges=["$A$1:$C$1"],
                )
            )

    def test_distinct_overlapping_merge_is_invalid(self):
        with self.assertRaisesRegex(ExcelFormulaError, INVALID_RESPONSE):
            parse_excel_table_result(
                table_payload(
                    ["A", None, None, None],
                    merges=["$A$1:$B$1", "$B$1:$B$2"],
                )
            )

    def test_malformed_dimensions_and_oversized_text_are_rejected(self):
        malformed = table_payload(
            ["A", "B", "C", "D"],
            address="$A$1:$C$2",
            rows=2,
            columns=2,
        )
        with self.assertRaisesRegex(ExcelFormulaError, INVALID_RESPONSE):
            parse_excel_table_result(malformed)

        oversized = table_payload(
            ["x" * 5_000_001, "B", "C", "D"],
        )
        with self.assertRaisesRegex(ExcelFormulaError, TOO_MUCH_TEXT):
            parse_excel_table_result(oversized)

        too_many_cells = table_payload(
            ["x"] * 10_001,
            address="$A$1:$A$10001",
            rows=10_001,
            columns=1,
        )
        with self.assertRaisesRegex(ExcelFormulaError, INVALID_RESPONSE):
            parse_excel_table_result(too_many_cells)

    def test_excel_not_running_and_executor_failure_are_content_free(self):
        with self.assertRaisesRegex(ExcelFormulaError, EXCEL_NOT_RUNNING):
            read_selected_excel_table(FakeExecutor(running=False))
        with self.assertRaisesRegex(ExcelFormulaError, EXECUTION_FAILED):
            read_selected_excel_table(
                FakeExecutor(error=RuntimeError("PRIVATE_CELL_CONTENT"))
            )

    def test_stable_reader_requires_two_equal_value_and_merge_snapshots(self):
        old = table_payload(["H", "V", "A", "1"])
        changed_value = table_payload(["H", "V", "A", "2"])
        changed_merge = table_payload(
            ["H", None, "A", "2"], merges=["$A$1:$B$1"]
        )

        for payloads, expected_reads in (
            ([old, old], 2),
            ([old, changed_value, changed_value], 3),
            ([changed_value, changed_merge, changed_merge], 3),
            ([["error", SELECTION_CHANGED], old, old], 3),
        ):
            with self.subTest(expected_reads=expected_reads, payloads=payloads):
                executor = FakeExecutor(payloads)
                selection = read_stable_selected_excel_table(executor)
                self.assertEqual(len(executor.sources), expected_reads)
                self.assertEqual(selection, parse_excel_table_result(payloads[-1]))

    def test_three_different_snapshots_fail_closed(self):
        executor = FakeExecutor(
            [
                table_payload(["H", "V", "A", "1"]),
                table_payload(["H", "V", "A", "2"]),
                table_payload(["H", "V", "A", "3"]),
            ]
        )

        with self.assertRaisesRegex(ExcelFormulaError, SELECTION_CHANGED):
            read_stable_selected_excel_table(executor)
        self.assertEqual(len(executor.sources), 3)

    def test_fatal_error_is_not_retried(self):
        for code in (EXCEL_NOT_RUNNING, DISPLAY_OVERFLOW):
            with self.subTest(code=code):
                executor = FakeExecutor([["error", code]])

                with self.assertRaisesRegex(ExcelFormulaError, code):
                    read_stable_selected_excel_table(executor)
                self.assertEqual(len(executor.sources), 1)


class ExcelTableHtmlBridgeTests(unittest.TestCase):
    @staticmethod
    def _selection(
        values,
        *,
        address,
        rows,
        columns,
        merges=(),
    ):
        return ExcelTableSelection(
            workbook="Book.xlsx",
            sheet="Sheet1",
            address=address,
            row_count=rows,
            column_count=columns,
            values=tuple(values),
            merge_areas=tuple(merges),
        )

    def test_multilevel_horizontal_and_vertical_merges_reach_existing_model(self):
        selection = self._selection(
            [
                "2026 실적",
                None,
                None,
                "직급",
                "1분기",
                None,
                None,
                "1월",
                "2월",
                "부장",
                "10",
                "20",
            ],
            address="$A$1:$C$4",
            rows=4,
            columns=3,
            merges=("$A$1:$C$1", "$A$2:$A$3", "$B$2:$C$2"),
        )

        model = excel_table_selection_to_model(selection)

        self.assertEqual(
            model,
            (
                [["직급", "1분기", "1분기"], ["직급", "1월", "2월"]],
                [["부장", "10", "20"]],
            ),
        )

    def test_one_column_selection_keeps_header_and_all_data_rows(self):
        selection = self._selection(
            ["이름", "사과", "배"],
            address="$A$1:$A$3",
            rows=3,
            columns=1,
        )

        model = excel_table_selection_to_model(selection)

        self.assertEqual(model, ([['이름']], [["사과"], ["배"]]))

    def test_explicit_trailing_blank_column_is_preserved(self):
        selection = self._selection(
            ["이름", "값", None, "사과", "2", None],
            address="$A$1:$C$2",
            rows=2,
            columns=3,
        )

        model = excel_table_selection_to_model(selection)

        self.assertEqual(model, ([['이름', '값', '']], [["사과", "2", ""]]))

    def test_cell_text_is_escaped_and_line_breaks_remain_inside_cell(self):
        selection = self._selection(
            ["이름", "값", "<&>", "첫째\r\n둘째"],
            address="$A$1:$B$2",
            rows=2,
            columns=2,
        )

        html = excel_table_selection_to_html(selection)
        model = html_table_to_model(html)

        self.assertIn("&lt;&amp;&gt;", html)
        self.assertIn("첫째<br>둘째", html)
        self.assertEqual(model[1][0], ["<&>", "첫째\n둘째"])

    def test_exact_data_whitespace_and_literal_br_survive_xml(self):
        selection = self._selection(
            [
                "이름",
                "표현",
                "줄바꿈",
                "  ACME   Inc.  ",
                "literal <br> and <br/>",
                "첫째\n둘째",
            ],
            address="$A$1:$C$2",
            rows=2,
            columns=3,
        )

        model = excel_table_selection_to_model(selection)
        xml = model_to_xml(*model)
        root = ET.fromstring(xml)
        values = [cell.text or "" for cell in root.findall(".//열")]

        self.assertEqual(
            values,
            ["  ACME   Inc.  ", "literal <br> and <br/>", "첫째\n둘째"],
        )
        self.assertIn("literal &lt;br&gt; and &lt;br/&gt;", xml)


if __name__ == "__main__":
    unittest.main()
