"""macOS direct Excel-selection table reader regression tests."""

from __future__ import annotations

from pathlib import Path
import unittest
import xml.etree.ElementTree as ET

from Foundation import NSAppleScript

from tablemark.converter.html_to_md import html_table_to_model
from tablemark.converter.table_xml import (
    TableXmlTooLargeError,
    model_to_xml,
    table_xml_to_model,
)
from tablemark.excel_formula import (
    _AE_LIST_DESCRIPTOR_TYPE,
    _descriptor_to_python,
    AUTOMATION_DENIED,
    NSAppleScriptExecutor,
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
    excel_table_selection_to_model_with_sources,
    excel_table_selection_xml_metadata,
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


class ExcelTableReadFailureClassificationTests(unittest.TestCase):
    @staticmethod
    def _script_result(*, read_error=None, capture_error=None,
                       verification_error=None, changed=None):
        # Execute the production AppleScript control flow with its Excel reads
        # replaced by deterministic inputs. No tell/application statement or
        # cell-reading handler remains, so this never contacts a running Excel.
        start = EXCEL_TABLE_SCRIPT.index("        try\n            set workbookName")
        end = EXCEL_TABLE_SCRIPT.index("\n    end tell\nend using terms from", start)
        body = EXCEL_TABLE_SCRIPT[start:end]
        read_start = body.index("            set compactValues")
        read_end = body.index("\n        on error", read_start)
        read_body = (
            f"            error number {read_error}"
            if read_error is not None else
            '            set valueKinds to "vvvv"\n'
            f'            set valueBlob to "H{_VALUE_DELIMITER}V{_VALUE_DELIMITER}A{_VALUE_DELIMITER}1"\n'
            '            set mergedAddresses to {}'
        )
        body = body[:read_start] + read_body + body[read_end:]
        replacements = {
            "set workbookName to name of selectedWorkbook as text": (
                f"error number {capture_error}" if capture_error is not None
                else 'set workbookName to "Book.xlsx"'
            ),
            "set sheetName to name of selectedSheet as text": 'set sheetName to "Sheet1"',
            "set selectionAddress to get address selectedRange row absolute true column absolute true reference style A1": 'set selectionAddress to "$A$1:$B$2"',
            "set currentWorkbookName to name of active workbook as text": (
                f"error number {verification_error}" if verification_error is not None
                else 'set currentWorkbookName to "' + ("Other.xlsx" if changed == "workbook" else "Book.xlsx") + '"'
            ),
            "set currentSheetName to name of active sheet as text": 'set currentSheetName to "' + ("Other" if changed == "sheet" else "Sheet1") + '"',
            "set currentRange to selection": 'set currentRange to "stub"',
            "set currentAreas to get areas of currentRange": "set currentAreas to " + ("{1, 2}" if changed == "areas" else "{1}"),
            "set currentAddress to get address currentRange row absolute true column absolute true reference style A1": 'set currentAddress to "' + ("$C$1:$D$2" if changed == "address" else "$A$1:$B$2") + '"',
            "count of rows of currentRange": "3" if changed == "rows" else "2",
            "count of columns of currentRange": "3" if changed == "columns" else "2",
        }
        for original, replacement in replacements.items():
            if body.count(original) != 1:
                raise AssertionError(f"production boundary changed: {original}")
            body = body.replace(original, replacement)
        if "tell application" in body or "using terms" in body:
            raise AssertionError("unexpected native application access")
        return NSAppleScriptExecutor().run(
            "set selectedRowCount to 2\nset selectedColumnCount to 2\n" + body
        )

    def test_normal_read_retains_payload_and_two_snapshot_success(self):
        payload = self._script_result()
        executor = FakeExecutor([payload, payload])

        selection = read_stable_selected_excel_table(executor)

        self.assertEqual(selection.values, ("H", "V", "A", "1"))
        self.assertEqual(len(executor.sources), 2)

    def test_failed_read_with_confirmed_identity_change_is_retryable(self):
        for changed in ("workbook", "sheet", "address", "areas", "rows", "columns"):
            with self.subTest(changed=changed):
                payload = self._script_result(read_error=-2700, changed=changed)
                self.assertEqual(payload, ["error", SELECTION_CHANGED])
                stable = self._script_result()
                executor = FakeExecutor([payload, stable, stable])

                selection = read_stable_selected_excel_table(executor)

                self.assertEqual(selection.values, ("H", "V", "A", "1"))
                self.assertEqual(len(executor.sources), 3)

    def test_failed_read_with_same_identity_fails_without_retry(self):
        payload = self._script_result(read_error=-2700)
        executor = FakeExecutor([payload])

        with self.assertRaisesRegex(ExcelFormulaError, EXECUTION_FAILED):
            read_stable_selected_excel_table(executor)

        self.assertEqual(len(executor.sources), 1)

    def test_unknown_identity_keeps_original_failure_classification(self):
        for arguments, expected in (
            ({"capture_error": -2700}, EXECUTION_FAILED),
            ({"read_error": -2700, "verification_error": -2700}, EXECUTION_FAILED),
            ({"verification_error": -2700}, SELECTION_CHANGED),
        ):
            with self.subTest(arguments=arguments):
                self.assertEqual(self._script_result(**arguments), ["error", expected])

    def test_automation_denial_remains_fatal_at_each_read_stage(self):
        for stage in ("capture_error", "read_error", "verification_error"):
            with self.subTest(stage=stage):
                with self.assertRaisesRegex(ExcelFormulaError, AUTOMATION_DENIED):
                    self._script_result(**{stage: -1743})
                executor = FakeExecutor(error=ExcelFormulaError(AUTOMATION_DENIED))
                with self.assertRaisesRegex(ExcelFormulaError, AUTOMATION_DENIED):
                    read_stable_selected_excel_table(executor)
                self.assertEqual(len(executor.sources), 1)

    def test_repeated_failed_reads_with_changed_identity_stop_after_three(self):
        payload = self._script_result(read_error=-2700, changed="address")
        executor = FakeExecutor([payload, payload, payload])

        with self.assertRaisesRegex(ExcelFormulaError, SELECTION_CHANGED):
            read_stable_selected_excel_table(executor)

        self.assertEqual(len(executor.sources), 3)


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

    def test_model_source_grid_tracks_original_cells_after_header_inference(self):
        selection = self._selection(
            [
                "그룹",
                "항목",
                "값",
                "A",
                "x",
                "1",
                None,
                "y",
                "2",
            ],
            address="$A$1:$C$3",
            rows=3,
            columns=3,
        )

        headers, rows, source_addresses = (
            excel_table_selection_to_model_with_sources(selection)
        )

        self.assertEqual(headers, [["그룹", "항목", "값"]])
        self.assertEqual(rows, [["A", "x", "1"], ["", "y", "2"]])
        self.assertEqual(
            source_addresses,
            [
                ["$A$2", "$B$2", "$C$2"],
                ["$A$3", "$B$3", "$C$3"],
            ],
        )

        titled = self._selection(
            [
                "2026 실적",
                None,
                None,
                "그룹",
                "항목",
                "값",
                "A",
                "x",
                "1",
                None,
                "y",
                "2",
            ],
            address="$A$1:$C$4",
            rows=4,
            columns=3,
            merges=("$A$1:$C$1",),
        )
        _, titled_rows, titled_sources = (
            excel_table_selection_to_model_with_sources(titled)
        )
        self.assertEqual(titled_rows, [["A", "x", "1"], ["", "y", "2"]])
        self.assertEqual(
            titled_sources,
            [
                ["$A$3", "$B$3", "$C$3"],
                ["$A$4", "$B$4", "$C$4"],
            ],
        )

    def test_xml_metadata_preserves_source_title_and_exact_merge_topology(self):
        selection = self._selection(
            [
                "2026 실적",
                None,
                None,
                "직급",
                "1분기",
                None,
                "부장",
                "10",
                "20",
            ],
            address="$A$1:$C$3",
            rows=3,
            columns=3,
            merges=("$A$1:$C$1", "$B$2:$C$2"),
        )
        headers, rows = excel_table_selection_to_model(selection)

        xml = model_to_xml(
            headers,
            rows,
            metadata=excel_table_selection_xml_metadata(selection, headers),
        )
        root = ET.fromstring(xml)

        self.assertEqual(root.attrib["형식버전"], "2")
        self.assertEqual(root.attrib["통합문서"], "Book.xlsx")
        self.assertEqual(root.attrib["시트"], "Sheet1")
        self.assertEqual(root.attrib["주소"], "$A$1:$C$3")
        self.assertEqual(root.attrib["헤더행수"], "1")
        self.assertEqual(
            root.attrib["병합범위"], "$A$1:$C$1 $B$2:$C$2"
        )
        self.assertEqual(root.attrib["제목수"], "1")
        self.assertEqual(root.attrib["제목1주소"], "$A$1:$C$1")
        self.assertEqual(root.attrib["제목1값"], "2026 실적")
        self.assertEqual(root.attrib["빈칸채움"], "미적용")
        self.assertEqual(root.attrib["빈칸채움수"], "0")
        self.assertEqual(root.attrib["빈칸채움셀"], "")

        repeated = self._selection(
            selection.values,
            address=selection.address,
            rows=selection.row_count,
            columns=selection.column_count,
        )
        repeated_headers, repeated_rows = excel_table_selection_to_model(repeated)
        repeated_xml = model_to_xml(
            repeated_headers,
            repeated_rows,
            metadata=excel_table_selection_xml_metadata(
                repeated, repeated_headers
            ),
        )
        self.assertNotEqual(xml, repeated_xml)
        self.assertEqual(ET.fromstring(repeated_xml).attrib["병합범위"], "")

        filled_xml = model_to_xml(
            headers,
            rows,
            metadata=excel_table_selection_xml_metadata(
                selection,
                headers,
                blank_fill_enabled=True,
                blank_fill_cells=("$A$3",),
            ),
        )
        filled_root = ET.fromstring(filled_xml)
        self.assertEqual(filled_root.attrib["빈칸채움"], "적용")
        self.assertEqual(
            filled_root.attrib["빈칸채움기준"],
            "왼쪽키열_위우선_좌측보완",
        )
        self.assertEqual(filled_root.attrib["빈칸채움수"], "1")
        self.assertEqual(filled_root.attrib["빈칸채움셀"], "$A$3")

    def test_blank_whitespace_and_duplicate_headers_roundtrip_exactly(self):
        headers = [["", " ID ", "값", "값"]]
        rows = [["A", "B", "10", "20"]]

        xml = model_to_xml(headers, rows)
        roundtrip = table_xml_to_model(xml)
        root = ET.fromstring(xml)
        cells = root.findall("행/열")

        self.assertEqual(roundtrip, (headers, rows))
        self.assertEqual(
            [(cell.attrib.get("i"), cell.attrib["n"]) for cell in cells],
            [("1", ""), (None, " ID "), ("3", "값"), ("4", "값")],
        )

    def test_multilevel_blank_leaf_does_not_inherit_group_name(self):
        headers = [["ID", "Q1", "Q1"], ["ID", "Jan", ""]]
        rows = [["A", "10", "20"]]

        xml = model_to_xml(headers, rows)
        roundtrip = table_xml_to_model(xml)
        leaf_names = [
            element.attrib["n"]
            for element in ET.fromstring(xml).findall(".//열")
        ]

        self.assertEqual(leaf_names, ["Jan", ""])
        self.assertEqual(roundtrip, (headers, rows))

    def test_blank_upper_header_keeps_excel_quarter_groups_and_metadata(self):
        selection = self._selection(
            [
                "항목", "", None, None, None,
                None, "1분기", None, "2분기", None,
                None, "매출", "비용", "매출", "비용",
                "제품A", "10", "5", "20", "8",
            ],
            address="$A$1:$E$4",
            rows=4,
            columns=5,
            merges=(
                "$A$1:$A$3", "$B$1:$E$1",
                "$B$2:$C$2", "$D$2:$E$2",
            ),
        )
        headers, rows = excel_table_selection_to_model(selection)
        metadata = excel_table_selection_xml_metadata(selection, headers)

        xml = model_to_xml(headers, rows, metadata=metadata)
        root = ET.fromstring(xml)
        row = root.find("행")

        self.assertEqual(row.attrib, {"항목": "제품A"})
        self.assertEqual(
            [group.attrib["이름"] for group in row.findall("열그룹")],
            ["1분기", "2분기"],
        )
        self.assertEqual(
            [(cell.attrib["i"], cell.attrib["n"], cell.text)
             for cell in row.findall("열그룹/열")],
            [("2", "매출", "10"), ("3", "비용", "5"),
             ("4", "매출", "20"), ("5", "비용", "8")],
        )
        self.assertEqual(root.attrib["헤더행수"], "3")
        self.assertEqual(root.attrib["병합범위"], " ".join(selection.merge_areas))
        self.assertEqual(root.attrib["빈칸채움수"], "0")
        # The unnamed upper level adds no semantic group; the source metadata
        # still records all three original header rows and the exact merges.
        roundtrip = table_xml_to_model(xml)
        self.assertEqual(roundtrip, (headers[1:], rows))
        self.assertEqual(model_to_xml(*roundtrip, metadata=metadata), xml)

    def test_blank_intermediate_header_keeps_nested_groups_and_exact_leaves(self):
        headers = [
            ["항목", "실적", "실적", "실적", "실적", "예산", "예산"],
            ["항목", "", "", " ", " ", "연간", "연간"],
            ["항목", "1분기", "1분기", "2분기", "2분기", "합계", "합계"],
            ["항목", "", "매출", "매출", " ", "계획", "실제"],
        ]
        rows = [["제품A", "10", "5", "20", "8", "50", "43"]]

        xml = model_to_xml(headers, rows)
        root = ET.fromstring(xml)
        actual = root.find("행/열그룹[@이름='실적']")

        self.assertEqual(
            [group.attrib["이름"] for group in actual.findall("열그룹")],
            ["1분기", "2분기"],
        )
        self.assertEqual(
            [(cell.attrib["i"], cell.attrib["n"])
             for cell in actual.findall("열그룹/열")],
            [("2", ""), ("3", "매출"), ("4", "매출"), ("5", " ")],
        )
        self.assertEqual(
            [cell.text for cell in root.findall(
                "행/열그룹[@이름='예산']/열그룹[@이름='연간']/"
                "열그룹[@이름='합계']/열"
            )],
            ["50", "43"],
        )
        roundtrip = table_xml_to_model(xml)
        self.assertEqual(roundtrip[1], rows)
        self.assertEqual(model_to_xml(*roundtrip), xml)

    def test_blank_upper_header_does_not_join_groups_across_named_neighbor(self):
        headers = [
            ["", "", "별도", "", ""],
            ["1분기", "1분기", "기타", "1분기", "1분기"],
            ["매출", "비용", "금액", "매출", "비용"],
        ]
        rows = [["10", "5", "7", "20", "8"]]

        xml = model_to_xml(headers, rows)
        groups = ET.fromstring(xml).findall("행/열그룹")

        self.assertEqual(
            [group.attrib["이름"] for group in groups],
            ["1분기", "별도", "1분기"],
        )
        self.assertEqual(
            [[cell.text for cell in group.iter("열")] for group in groups],
            [["10", "5"], ["7"], ["20", "8"]],
        )
        self.assertEqual(model_to_xml(*table_xml_to_model(xml)), xml)

    def test_blank_upper_header_with_only_leaves_keeps_existing_xml(self):
        headers = [["", " "], ["ID", "값"], ["ID", "값"]]
        rows = [["제품A", "10"]]

        self.assertEqual(
            model_to_xml(headers, rows),
            '<표>\n  <행>\n    <열 n="ID">제품A</열>\n'
            '    <열 n="값">10</열>\n  </행>\n</표>',
        )

    def test_title_metadata_matches_one_column_and_multirow_merge_rules(self):
        one_column = self._selection(
            ["그룹", None, "값"],
            address="$A$1:$A$3",
            rows=3,
            columns=1,
            merges=("$A$1:$A$2",),
        )
        one_headers, one_rows = excel_table_selection_to_model(one_column)
        one_metadata = excel_table_selection_xml_metadata(
            one_column, one_headers
        )

        self.assertEqual(one_metadata.title_rows, ())
        self.assertEqual(one_headers, [["그룹"]])
        self.assertEqual(one_rows, [["그룹"], ["값"]])

        multirow_title = self._selection(
            [
                "2026 실적",
                None,
                None,
                None,
                None,
                None,
                "항목",
                "Q1",
                "Q2",
                "매출",
                "10",
                "20",
            ],
            address="$A$1:$C$4",
            rows=4,
            columns=3,
            merges=("$A$1:$C$2",),
        )
        title_headers, title_rows = excel_table_selection_to_model(
            multirow_title
        )
        title_metadata = excel_table_selection_xml_metadata(
            multirow_title, title_headers
        )

        self.assertEqual(title_headers, [["항목", "Q1", "Q2"]])
        self.assertEqual(title_rows, [["매출", "10", "20"]])
        self.assertEqual(
            [
                (title.address, title.value)
                for title in title_metadata.title_rows
            ],
            [("$A$1:$C$2", "2026 실적")],
        )

    def test_general_xml_byte_limit_stops_during_serialization(self):
        headers = [["H" * 1000]]
        rows = [[""] for _ in range(100)]

        with self.assertRaises(TableXmlTooLargeError):
            model_to_xml(headers, rows, max_bytes=5_000)

        xml = model_to_xml([["한글"]], [["<&>"]], max_bytes=10_000)
        encoded_size = len(xml.encode("utf-8"))
        self.assertEqual(
            model_to_xml(
                [["한글"]],
                [["<&>"]],
                max_bytes=encoded_size,
            ),
            xml,
        )
        with self.assertRaises(TableXmlTooLargeError):
            model_to_xml(
                [["한글"]],
                [["<&>"]],
                max_bytes=encoded_size - 1,
            )


if __name__ == "__main__":
    unittest.main()
