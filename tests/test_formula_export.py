from __future__ import annotations

from dataclasses import replace
import unittest
import xml.etree.ElementTree as ET
from unittest.mock import patch

from tablemark.converter.formula_export import (
    ExcelFormulaReference,
    ExcelFormulaSelection,
    ExcelReferenceCell,
    ExcelSelectionCell,
    FormulaXmlTooLargeError,
    analyze_formula_references,
    extract_formula_reference_targets,
    formula_copy_notice_key,
    formula_selection_to_xml,
    normalize_excel_number_text,
)


class FormulaExportTests(unittest.TestCase):
    def test_expands_native_numeric_exponents_but_not_literal_text(self):
        self.assertEqual(normalize_excel_number_text("3.15E+5"), "315000")
        self.assertEqual(normalize_excel_number_text(" 1.2e-7"), "0.00000012")
        self.assertEqual(normalize_excel_number_text("-0.0"), "0")
        # Type discrimination happens in the Excel adapters. A text cell is
        # never passed to the numeric normalizer and therefore stays literal.
        text_cell = ExcelReferenceCell("$A$1", "3.15E+5", "text")
        number_cell = ExcelReferenceCell("$B$1", "3.15E+5", "number")
        self.assertEqual(text_cell.value, "3.15E+5")
        self.assertEqual(normalize_excel_number_text(number_cell.value), "315000")

    def test_serializes_same_and_cross_sheet_reference_values_under_formula(self):
        selection = self._one_cell(value="120", formula_a1="=C6*'단가 표'!D6")
        cell = selection.cells[0]
        selection = ExcelFormulaSelection(
            selection.workbook,
            selection.sheet,
            selection.address,
            selection.row_count,
            selection.column_count,
            (
                ExcelSelectionCell(
                    cell.address,
                    cell.value,
                    cell.formula_a1,
                    cell.formula_r1c1,
                    references=(
                        ExcelFormulaReference(
                            "Sheet1",
                            "$C$6",
                            (ExcelReferenceCell("$C$6", "10"),),
                        ),
                        ExcelFormulaReference(
                            "단가 표",
                            "$D$6",
                            (ExcelReferenceCell("$D$6", "12"),),
                        ),
                    ),
                ),
            ),
        )

        xml = formula_selection_to_xml(selection)

        root = ET.fromstring(xml)
        formula_cell = root.find("행/셀")
        self.assertEqual(formula_cell.attrib["값"], "120")
        references = formula_cell.findall("참조범위")
        self.assertEqual(
            [(item.attrib["시트"], item.attrib["주소"]) for item in references],
            [("Sheet1", "$C$6"), ("단가 표", "$D$6")],
        )
        self.assertEqual(
            [item.find("참조셀").attrib["값"] for item in references],
            ["10", "12"],
        )

    def test_serializes_formula_result_and_substituted_range_values(self):
        selection = ExcelFormulaSelection(
            "Tabledown.xlsx",
            "수식_스냅샷",
            "$F$10",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$F$10",
                    "862300",
                    "=SUM(F4:F9)",
                    "=SUM(R[-6]C:R[-1]C)",
                    references=(
                        ExcelFormulaReference(
                            "수식_스냅샷",
                            "$F$4:$F$9",
                            tuple(
                                ExcelReferenceCell(address, value, "number")
                                for address, value in (
                                    ("$F$4", "75000"),
                                    ("$F$5", "136800"),
                                    ("$F$6", "315000"),
                                    ("$F$7", "72000"),
                                    ("$F$8", "263500"),
                                    ("$F$9", "0"),
                                )
                            ),
                        ),
                    ),
                ),
            ),
        )

        formula_cell = ET.fromstring(formula_selection_to_xml(selection)).find(
            "행/셀"
        )

        self.assertEqual(formula_cell.attrib["값"], "862300")
        self.assertEqual(
            formula_cell.attrib["값대입수식"],
            "=SUM({75000;136800;315000;72000;263500;0})",
        )

    def test_substitutes_repeated_cross_sheet_and_two_dimensional_references(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$C$3",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$C$3",
                    "ok",
                    "=A1+A1&'Data Sheet'!B2&SUM(D4:E5)",
                    references=(
                        ExcelFormulaReference(
                            "Sheet1",
                            "$A$1",
                            (ExcelReferenceCell("$A$1", "2", "number"),),
                        ),
                        ExcelFormulaReference(
                            "Data Sheet",
                            "$B$2",
                            (ExcelReferenceCell("$B$2", 'a"b', "text"),),
                        ),
                        ExcelFormulaReference(
                            "Sheet1",
                            "$D$4:$E$5",
                            (
                                ExcelReferenceCell("$D$4", "1", "number"),
                                ExcelReferenceCell("$E$4", "true", "boolean"),
                                ExcelReferenceCell("$D$5", None, "blank"),
                                ExcelReferenceCell("$E$5", "#N/A", "error"),
                            ),
                        ),
                    ),
                ),
            ),
        )

        cell = ET.fromstring(formula_selection_to_xml(selection)).find("행/셀")

        self.assertEqual(
            cell.attrib["값대입수식"],
            '=2+2&"a""b"&SUM({1,TRUE;BLANK(),#N/A})',
        )

    def test_substituted_formula_is_omitted_for_partial_or_untyped_references(self):
        cases = (
            ExcelSelectionCell(
                "$B$1",
                "1",
                '=INDIRECT("A1")',
                references_complete=False,
            ),
            ExcelSelectionCell(
                "$B$1",
                "1",
                "=A1",
                references=(
                    ExcelFormulaReference(
                        "Sheet1", "$A$1", (ExcelReferenceCell("$A$1", "1"),)
                    ),
                ),
            ),
        )
        for formula_cell in cases:
            with self.subTest(formula=formula_cell.formula_a1):
                selection = ExcelFormulaSelection(
                    "Book.xlsx", "Sheet1", "$B$1", 1, 1, (formula_cell,)
                )
                cell = ET.fromstring(formula_selection_to_xml(selection)).find(
                    "행/셀"
                )
                self.assertNotIn("값대입수식", cell.attrib)

    def test_substituted_formula_size_cap_never_blocks_original_export(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$B$1",
                    "1",
                    "=A1",
                    references=(
                        ExcelFormulaReference(
                            "Sheet1",
                            "$A$1",
                            (ExcelReferenceCell("$A$1", "123", "number"),),
                        ),
                    ),
                ),
            ),
        )

        with patch(
            "tablemark.converter.formula_export.MAX_SUBSTITUTED_FORMULA_CHARACTERS",
            2,
        ):
            cell = ET.fromstring(formula_selection_to_xml(selection)).find("행/셀")

        self.assertEqual(cell.attrib["수식"], "=A1")
        self.assertNotIn("값대입수식", cell.attrib)
        self.assertEqual(
            cell.attrib["값대입수식상태"],
            "omitted_character_limit",
        )

    def test_character_cap_does_not_mark_formula_without_direct_references(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1",
            1,
            1,
            (ExcelSelectionCell("$A$1", "1", "=ROW()"),),
        )

        with patch(
            "tablemark.converter.formula_export.MAX_SUBSTITUTED_FORMULA_CHARACTERS",
            0,
        ):
            cell = ET.fromstring(formula_selection_to_xml(selection)).find("행/셀")

        self.assertNotIn("값대입수식", cell.attrib)
        self.assertNotIn("값대입수식상태", cell.attrib)

    def test_xml_byte_cap_drops_substitution_before_failing_export(self):
        long_reference_value = "x" * 1_000
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$B$1",
                    "123",
                    "=A1",
                    references=(
                        ExcelFormulaReference(
                            "Sheet1",
                            "$A$1",
                            (
                                ExcelReferenceCell(
                                    "$A$1",
                                    long_reference_value,
                                    "text",
                                ),
                            ),
                        ),
                    ),
                ),
            ),
        )
        with patch(
            "tablemark.converter.formula_export.MAX_SUBSTITUTED_FORMULA_CHARACTERS",
            0,
        ):
            base_xml = formula_selection_to_xml(selection)

        with patch(
            "tablemark.converter.formula_export.MAX_XML_BYTES",
            len(base_xml.encode("utf-8")) + 100,
        ):
            xml = formula_selection_to_xml(selection)

        root = ET.fromstring(xml)
        self.assertEqual(
            root.attrib["값대입수식상태"],
            "omitted_xml_size_limit",
        )
        self.assertNotIn("값대입수식", root.find("행/셀").attrib)
        self.assertEqual(root.attrib["계산결과상태"], "freshness_unverified")

    def test_xml_byte_cap_never_silently_drops_explanatory_metadata(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$B$1",
                    "1",
                    '=INDIRECT("A1")',
                    references_complete=False,
                    value_kind="number",
                    reference_issues=("dynamic_reference",),
                ),
            ),
            calculation_mode="manual",
            calculation_state="pending",
        )
        with patch(
            "tablemark.converter.formula_export.MAX_SUBSTITUTED_FORMULA_CHARACTERS",
            0,
        ):
            complete_xml = formula_selection_to_xml(selection)

        with patch(
            "tablemark.converter.formula_export.MAX_XML_BYTES",
            len(complete_xml.encode("utf-8")) - 1,
        ), self.assertRaises(FormulaXmlTooLargeError):
            formula_selection_to_xml(selection)

    def test_shared_reference_rendering_stops_at_byte_limit(self):
        reference = ExcelFormulaReference(
            "Sheet1",
            "$Z$1:$Z$2048",
            tuple(
                ExcelReferenceCell(
                    f"$Z${row}",
                    "x" * 100,
                    "text",
                )
                for row in range(1, 2049)
            ),
        )
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1:$A$300",
            300,
            1,
            tuple(
                ExcelSelectionCell(
                    f"$A${row}",
                    "1",
                    "=SUM($Z$1:$Z$2048)",
                    references=(reference,),
                    value_kind="number",
                )
                for row in range(1, 301)
            ),
        )

        with patch(
            "tablemark.converter.formula_export.MAX_XML_BYTES", 100_000
        ), self.assertRaises(FormulaXmlTooLargeError):
            formula_selection_to_xml(selection)

    def test_partial_static_reference_state_is_explicit(self):
        selection = self._one_cell(value="10", formula_a1='=INDIRECT("A1")')
        cell = selection.cells[0]
        selection = ExcelFormulaSelection(
            selection.workbook,
            selection.sheet,
            selection.address,
            1,
            1,
            (
                ExcelSelectionCell(
                    cell.address,
                    cell.value,
                    cell.formula_a1,
                    references_complete=False,
                ),
            ),
        )

        formula_cell = ET.fromstring(formula_selection_to_xml(selection)).find(
            "행/셀"
        )

        self.assertEqual(formula_cell.attrib["참조상태"], "일부")

    def test_extracts_static_a1_ranges_and_cross_sheet_references(self):
        targets, complete = extract_formula_reference_targets(
            "=C6*D6+SUM('단가 표'!B2:B4)", "기본_표"
        )

        self.assertTrue(complete)
        self.assertEqual(
            [(item.sheet, item.address, item.row_count, item.column_count) for item in targets],
            [
                ("기본_표", "$C$6", 1, 1),
                ("기본_표", "$D$6", 1, 1),
                ("단가 표", "$B$2:$B$4", 3, 1),
            ],
        )

    def test_reference_extraction_ignores_strings_and_marks_unsupported_partial(self):
        static_targets, static_complete = extract_formula_reference_targets(
            '=IF(A1="Sheet2!B2",Sheet2!C3,0)', "Sheet1"
        )
        dynamic_targets, dynamic_complete = extract_formula_reference_targets(
            '=INDIRECT("Sheet2!A1")+Table1[A1]+B1', "Sheet1"
        )
        external_targets, external_complete = extract_formula_reference_targets(
            "='[Other.xlsx]Sheet1'!A1", "Sheet1"
        )

        self.assertTrue(static_complete)
        self.assertEqual(
            [(item.sheet, item.address) for item in static_targets],
            [("Sheet1", "$A$1"), ("Sheet2", "$C$3")],
        )
        self.assertFalse(dynamic_complete)
        self.assertEqual(
            [(item.sheet, item.address) for item in dynamic_targets],
            [("Sheet1", "$B$1")],
        )
        self.assertFalse(external_complete)
        self.assertEqual(external_targets, ())

    def test_callable_names_do_not_make_builtin_functions_false_partial(self):
        for formula in (
            "=SUM(A1:A2)",
            "=_xlfn.SEQUENCE(A1,2)",
            "=ROUND(A1,0)",
            "=MEDIAN(A1:A2)",
            "=STDEV.S(A1:A2)",
            "=NETWORKDAYS(A1,A2)",
            "=FORECAST.LINEAR(A1,A2:A3,B2:B3)",
            "=TAKE(A1:B2,1)",
        ):
            with self.subTest(formula=formula):
                _targets, complete = extract_formula_reference_targets(
                    formula, "Sheet1"
                )
                self.assertTrue(complete)

    def test_unicode_defined_names_are_not_read_as_embedded_a1_references(self):
        for name in ("매출A1", "A1매출", "éA1", "A1é", "𐐀A1", "\\A1"):
            with self.subTest(name=name):
                targets, issues = analyze_formula_references(
                    f"=SUM({name})+$B$2", "Sheet1"
                )

                self.assertEqual(
                    [(target.sheet, target.address) for target in targets],
                    [("Sheet1", "$B$2")],
                )
                self.assertEqual(issues, ("defined_name_or_unsupported_syntax",))

    def test_unicode_sheet_names_and_normal_index_lookup_remain_supported(self):
        targets, issues = analyze_formula_references(
            "=매출!$A$1+SUM('단가 표'!B$2:$C3)+INDEX(D1:E3,2,1)",
            "Sheet1",
        )

        self.assertEqual(
            [(target.sheet, target.address) for target in targets],
            [
                ("매출", "$A$1"),
                ("단가 표", "$B$2:$C$3"),
                ("Sheet1", "$D$1:$E$3"),
            ],
        )
        self.assertEqual(issues, ())
        string_targets, string_issues = analyze_formula_references(
            '=IF(A1="A1:INDEX(B1:B10,5)",B2,0)', "Sheet1"
        )
        self.assertEqual(
            [target.address for target in string_targets], ["$A$1", "$B$2"]
        )
        self.assertEqual(string_issues, ())

    def test_calculated_range_endpoints_are_partial_without_guessing_cells(self):
        for formula in (
            "=SUM(A1:INDEX(B1:B10,5))",
            "=SUM(A1 : (INDEX(B1:B10,5)))",
            "=SUM(INDEX(B1:B10,5):A1)",
            "=SUM(A1:_xlfn.INDEX(B1:B10,5))",
            "=SUM(A1:CHOOSE(1,B5,B10))",
        ):
            with self.subTest(formula=formula):
                targets, issues = analyze_formula_references(formula, "Sheet1")

                self.assertIn("calculated_range_reference", issues)
                self.assertNotIn("$A$1:$B$5", [target.address for target in targets])

        targets, issues = analyze_formula_references(
            "=SUM(A1:INDEX(B1:B10,5))", "Sheet1"
        )
        self.assertEqual(
            [target.address for target in targets], ["$A$1", "$B$1:$B$10"]
        )
        selection = self._one_cell(
            value="30", formula_a1="=SUM(A1:INDEX(B1:B10,5))"
        )
        selection = replace(
            selection,
            address="$D$1",
            cells=(replace(
                selection.cells[0],
                address="$D$1",
                references=(
                    ExcelFormulaReference(
                        "Sheet1", "$A$1",
                        (ExcelReferenceCell("$A$1", "1", "number"),),
                    ),
                    ExcelFormulaReference(
                        "Sheet1", "$B$1:$B$10",
                        tuple(
                            ExcelReferenceCell(f"$B${row}", str(row), "number")
                            for row in range(1, 11)
                        ),
                    ),
                ),
                references_complete=not issues,
                reference_issues=issues,
            ),),
        )
        exported = ET.fromstring(formula_selection_to_xml(selection)).find("행/셀")
        self.assertEqual(exported.attrib["값"], "30")
        self.assertEqual(exported.attrib["수식"], "=SUM(A1:INDEX(B1:B10,5))")
        self.assertEqual(exported.attrib["참조상태"], "일부")
        self.assertIn("calculated_range_reference", exported.attrib["참조누락이유"])
        self.assertNotIn("값대입수식", exported.attrib)

    def test_broken_sheet_references_do_not_read_a_real_ref_worksheet(self):
        for broken in ("#REF!B2", "#REF!$B$2", "#REF!B2:C3"):
            with self.subTest(broken=broken):
                targets, issues = analyze_formula_references(
                    f"=IFERROR({broken},A1)", "Sheet1"
                )

                self.assertEqual(
                    [(target.sheet, target.address) for target in targets],
                    [("Sheet1", "$A$1")],
                )
                self.assertEqual(issues, ("invalid_a1_reference",))

        for sheet in ("#REF!", "예산 #REF!"):
            with self.subTest(sheet=sheet):
                targets, issues = analyze_formula_references(
                    f"='{sheet}'!A1", "Sheet1"
                )
                self.assertEqual(
                    [(target.sheet, target.address) for target in targets],
                    [(sheet, "$A$1")],
                )
                self.assertEqual(issues, ())

    def test_parenthesized_and_error_range_endpoints_are_partial(self):
        for formula in (
            "=SUM(A1:(B2))",
            "=SUM(#REF!:$A$1)",
            "=SUM(#REF!: $A$1)",
            "=SUM(A1 : (B2))",
            "=SUM((A1):B2)",
            "=SUM(A1:#REF!)",
            "=SUM(A1 : #REF!)",
            "=SUM(#REF! : A1)",
            "=SUM(A1:B2:(C3))",
        ):
            with self.subTest(formula=formula):
                targets, issues = analyze_formula_references(formula, "Sheet1")

                self.assertIn("calculated_range_reference", issues)
                self.assertNotIn("$A$1:$C$3", [target.address for target in targets])
                selection = self._one_cell(value="10", formula_a1=formula)
                selection = replace(
                    selection,
                    cells=(replace(
                        selection.cells[0],
                        references_complete=not issues,
                        reference_issues=issues,
                    ),),
                )
                cell = ET.fromstring(formula_selection_to_xml(selection)).find("행/셀")
                self.assertEqual(cell.attrib["수식"], formula)
                self.assertEqual(cell.attrib["참조상태"], "일부")
                self.assertNotIn("값대입수식", cell.attrib)

    def test_static_ranges_and_colons_inside_strings_keep_existing_status(self):
        for formula in (
            "=(A1)",
            "=SUM((A1:B2))",
            "=SUM(A1 : B2)",
            "=SUM($A$1:\t$B$2)",
            "=INDEX(A1:B2,2,2)",
            "=SUM('입력 표'!A1:B2)",
            "=SUM(A1:B2,C3:D4)",
            '=IF(A1=":",B2,C3)',
            '=HYPERLINK("https://example.test",A1)',
            "=IFERROR(A1,#REF!)",
        ):
            with self.subTest(formula=formula):
                _targets, issues = analyze_formula_references(formula, "Sheet1")
                self.assertEqual(issues, ())

        for formula in ("=SUM(A:A)+B2", "=SUM(1:1)+B2"):
            with self.subTest(formula=formula):
                _targets, issues = analyze_formula_references(formula, "Sheet1")
                self.assertIn("whole_row_or_column_reference", issues)
                self.assertNotIn("calculated_range_reference", issues)

    def test_three_dimensional_references_are_not_misread_as_current_sheet(self):
        for formula in (
            "=SUM(Sheet1:Sheet3!A1)",
            "=SUM('Sheet 1':'Sheet 3'!A1:B2)",
            "=SUM(Sheet1:Sheet3!A1 : B2)",
            "=SUM('Sheet 1':'Sheet 3'!A1\t:\tB2)",
            "=SUM(Sheet1:Sheet3!A1\n:\nB2)",
        ):
            with self.subTest(formula=formula):
                targets, complete = extract_formula_reference_targets(
                    formula, "Current"
                )

                self.assertEqual(targets, ())
                self.assertFalse(complete)

        mixed_targets, mixed_complete = extract_formula_reference_targets(
            "=B1+SUM(Sheet1:Sheet3!A1)", "Current"
        )
        self.assertEqual(
            [(item.sheet, item.address) for item in mixed_targets],
            [("Current", "$B$1")],
        )
        self.assertFalse(mixed_complete)

    def test_numeric_exponent_constants_do_not_make_static_references_partial(self):
        for formula in ("=A1+1E+5", "=A1+1.2E-7", "=A1+.5E2"):
            with self.subTest(formula=formula):
                targets, complete = extract_formula_reference_targets(
                    formula, "Sheet1"
                )

                self.assertTrue(complete)
                self.assertEqual(
                    [(target.sheet, target.address) for target in targets],
                    [("Sheet1", "$A$1")],
                )

    def test_excel_error_literals_do_not_create_false_partial_references(self):
        for formula in (
            "=IFERROR(A1,#N/A)",
            "=IF(A1=0,#DIV/0!,A1)",
            "=IF(A1=0,#NAME?,A1)",
            "=IFERROR(A1,#REF!)",
            "=IFERROR(#N/A/A1,0)",
            "=IFERROR(#GETTING_DATA/A1,0)",
            "=IFERROR(#DIV/0!/A1,0)",
        ):
            with self.subTest(formula=formula):
                targets, issues = analyze_formula_references(formula, "Sheet1")

                self.assertEqual(
                    [(target.sheet, target.address) for target in targets],
                    [("Sheet1", "$A$1")],
                )
                self.assertEqual(issues, ())

    def test_spaced_ranges_are_supported_and_whole_rows_are_partial(self):
        targets, complete = extract_formula_reference_targets(
            "=SUM(A1 : B2)", "Sheet1"
        )
        multiline_targets, multiline_complete = extract_formula_reference_targets(
            "=SUM(A1\n:\nB2)", "Sheet1"
        )
        whole_row_targets, whole_row_complete = extract_formula_reference_targets(
            "=SUM(1:1)", "Sheet1"
        )

        self.assertTrue(complete)
        self.assertEqual(
            [(target.sheet, target.address) for target in targets],
            [("Sheet1", "$A$1:$B$2")],
        )
        self.assertTrue(multiline_complete)
        self.assertEqual(
            [(target.sheet, target.address) for target in multiline_targets],
            [("Sheet1", "$A$1:$B$2")],
        )
        self.assertEqual(whole_row_targets, ())
        self.assertFalse(whole_row_complete)

    def test_serializes_full_table_row_major_deterministically(self):
        selection = ExcelFormulaSelection(
            workbook="매출.xlsx",
            sheet="요약",
            address="A1:C2",
            row_count=2,
            column_count=3,
            cells=(
                ExcelSelectionCell("A1", "10"),
                ExcelSelectionCell("B1", "20"),
                ExcelSelectionCell("C1", "30", "=A1+B1", "=RC[-2]+RC[-1]"),
                ExcelSelectionCell("A2", None),
                ExcelSelectionCell("B2", ""),
                ExcelSelectionCell("C2", "정상", '=IF(A2="","정상","확인")'),
            ),
        )

        self.assertEqual(
            formula_selection_to_xml(selection),
            '<표범위 통합문서="매출.xlsx" 시트="요약" 주소="A1:C2" 행수="2" 열수="3" 형식버전="2" 계산모드="unknown" 계산상태="unavailable" 계산결과상태="freshness_unverified" 값기준="Excel현재원시값" 표시정보상태="미포함" 병합정보상태="미포함">\n'
            '  <행 인덱스="1">\n'
            '    <셀 주소="A1" 값="10" />\n'
            '    <셀 주소="B1" 값="20" />\n'
            '    <셀 주소="C1" 값="30" 수식="=A1+B1" 수식R1C1="=RC[-2]+RC[-1]" />\n'
            '  </행>\n'
            '  <행 인덱스="2">\n'
            '    <셀 주소="A2" />\n'
            '    <셀 주소="B2" 값="" />\n'
            '    <셀 주소="C2" 값="정상" 수식="=IF(A2=&quot;&quot;,&quot;정상&quot;,&quot;확인&quot;)" />\n'
            '  </행>\n'
            '</표범위>',
        )

    def test_serializes_value_kinds_calculation_state_and_partial_reason(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1:$B$1",
            1,
            2,
            (
                ExcelSelectionCell(
                    "$A$1",
                    "123",
                    value_kind="number",
                ),
                ExcelSelectionCell(
                    "$B$1",
                    "2",
                    '=INDIRECT("A1")*2',
                    references_complete=False,
                    value_kind="text",
                    reference_issues=("dynamic_reference",),
                ),
            ),
            calculation_mode="manual",
            calculation_state="pending",
        )

        root = ET.fromstring(formula_selection_to_xml(selection))
        cells = root.findall("행/셀")

        self.assertEqual(root.attrib["형식버전"], "2")
        self.assertEqual(root.attrib["계산모드"], "manual")
        self.assertEqual(root.attrib["계산상태"], "pending")
        self.assertEqual(root.attrib["계산결과상태"], "calculation_incomplete")
        self.assertEqual(cells[0].attrib["값종류"], "number")
        self.assertEqual(cells[1].attrib["값종류"], "text")
        self.assertEqual(cells[1].attrib["참조상태"], "일부")
        self.assertEqual(
            cells[1].attrib["참조누락이유"], "dynamic_reference"
        )

    def test_formula_copy_notice_combines_partial_references_and_calculation(self):
        selection = self._one_cell(value="2", formula_a1="=1+1")
        for calculation_state in (
            "done", "calculating", "pending", "unknown", "unavailable", None
        ):
            for partial in (False, True):
                with self.subTest(state=calculation_state, partial=partial):
                    current = replace(
                        selection,
                        address="$A$1:$B$1",
                        column_count=2,
                        cells=(
                            ExcelSelectionCell("$A$1", "label"),
                            replace(
                                selection.cells[0],
                                address="$B$1",
                                references_complete=not partial,
                            ),
                        ),
                        calculation_state=calculation_state,
                    )
                    calculating = calculation_state in {"calculating", "pending"}
                    expected = {
                        (False, False): None,
                        (True, False): "partial_references",
                        (False, True): "calculation_incomplete",
                        (True, True): "partial_references_and_calculation",
                    }[(partial, calculating)]
                    self.assertEqual(formula_copy_notice_key(current), expected)

    def test_substituted_formula_is_self_described_as_non_equivalent(self):
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$B$1",
            1,
            1,
            (
                ExcelSelectionCell(
                    "$B$1",
                    "1",
                    "=ROW(A1)",
                    references=(
                        ExcelFormulaReference(
                            "Sheet1",
                            "$A$1",
                            (ExcelReferenceCell("$A$1", "10", "number"),),
                        ),
                    ),
                    value_kind="number",
                ),
            ),
        )

        cell = ET.fromstring(formula_selection_to_xml(selection)).find("행/셀")

        self.assertEqual(cell.attrib["값대입수식"], "=ROW(10)")
        self.assertEqual(cell.attrib["값대입수식동등성"], "보장안함")
        self.assertEqual(
            cell.find("참조범위/참조셀").attrib["값종류"], "number"
        )

    def test_shared_reference_counts_once_for_serializer_limits(self):
        reference = ExcelFormulaReference(
            "Sheet1",
            "$Z$1",
            (ExcelReferenceCell("$Z$1", "1", "number"),),
        )
        cells = tuple(
            ExcelSelectionCell(
                f"$A${index}",
                "1",
                "=$Z$1",
                references=(reference,),
                value_kind="number",
            )
            for index in range(1, 301)
        )
        selection = ExcelFormulaSelection(
            "Book.xlsx",
            "Sheet1",
            "$A$1:$A$300",
            300,
            1,
            cells,
        )

        root = ET.fromstring(formula_selection_to_xml(selection))

        self.assertEqual(len(root.findall(".//참조범위")), 300)
        self.assertFalse(any(
            cell.attrib.get("참조상태") == "일부"
            for cell in root.findall(".//셀")
        ))

    def test_reference_analysis_reports_actionable_reasons(self):
        targets, issues = analyze_formula_references(
            '=INDIRECT("A1")+SUM(A:A)+MyName', "Sheet1"
        )

        self.assertEqual(targets, ())
        self.assertEqual(
            issues,
            (
                "dynamic_reference",
                "whole_row_or_column_reference",
                "defined_name_or_unsupported_syntax",
            ),
        )

    def test_escapes_metadata_address_value_and_formula_attributes(self):
        selection = ExcelFormulaSelection(
            workbook='Q&A "최종".xlsx',
            sheet="<입력>",
            address="A1&A2",
            row_count=1,
            column_count=1,
            cells=(ExcelSelectionCell('A1"&', '<결과 "1">', '="<tag>"&A1'),),
        )

        xml = formula_selection_to_xml(selection)

        self.assertIn('통합문서="Q&amp;A &quot;최종&quot;.xlsx"', xml)
        self.assertIn('시트="&lt;입력&gt;"', xml)
        self.assertIn('주소="A1&amp;A2"', xml)
        self.assertIn('<셀 주소="A1&quot;&amp;"', xml)
        self.assertIn('값="&lt;결과 &quot;1&quot;&gt;"', xml)
        self.assertIn('수식="=&quot;&lt;tag&gt;&quot;&amp;A1"', xml)
        cell = ET.fromstring(xml).find("행/셀")
        self.assertEqual(cell.attrib["값"], '<결과 "1">')
        self.assertEqual(cell.attrib["수식"], '="<tag>"&A1')

    def test_preserves_control_whitespace_without_formula_injection_lines(self):
        formula = '="safe"\t=HYPERLINK("https://example.invalid")\n=1+1\r=2+2'
        value = "첫째\t둘째\n셋째\r넷째"
        selection = self._one_cell(value=value, formula_a1=formula, formula_r1c1=formula)

        xml = formula_selection_to_xml(selection)

        self.assertIn("&#09;", xml)
        self.assertIn("&#10;", xml)
        self.assertIn("&#13;", xml)
        self.assertFalse(any(line.lstrip().startswith("=") for line in xml.splitlines()))
        cell = ET.fromstring(xml).find("행/셀")
        self.assertEqual(cell.attrib["값"], value)
        self.assertEqual(cell.attrib["수식"], formula)
        self.assertEqual(cell.attrib["수식R1C1"], formula)

    def test_formula_result_value_and_missing_r1c1_are_supported(self):
        xml = formula_selection_to_xml(self._one_cell(value="2", formula_a1="=1+1"))

        cell = ET.fromstring(xml).find("행/셀")
        self.assertEqual(cell.attrib["값"], "2")
        self.assertEqual(cell.attrib["수식"], "=1+1")
        self.assertNotIn("수식R1C1", cell.attrib)

    def test_r1c1_only_formula_counts_as_a_formula(self):
        xml = formula_selection_to_xml(
            self._one_cell(value="2", formula_a1=None, formula_r1c1="=1+1")
        )

        cell = ET.fromstring(xml).find("행/셀")
        self.assertNotIn("수식", cell.attrib)
        self.assertEqual(cell.attrib["수식R1C1"], "=1+1")

    def test_rejects_too_few_or_too_many_cells_for_shape(self):
        for cells in (
            (ExcelSelectionCell("A1", None, "=1"),),
            (
                ExcelSelectionCell("A1", None, "=1"),
                ExcelSelectionCell("B1", None),
                ExcelSelectionCell("C1", None),
            ),
        ):
            with self.subTest(cell_count=len(cells)):
                selection = ExcelFormulaSelection(
                    "Book1.xlsx", "Sheet1", "A1:B1", 1, 2, cells
                )
                with self.assertRaisesRegex(ValueError, "셀 수가 일치하지 않습니다"):
                    formula_selection_to_xml(selection)

    def test_rejects_non_positive_dimensions(self):
        for row_count, column_count in ((0, 1), (1, 0), (-1, 1)):
            with self.subTest(row_count=row_count, column_count=column_count):
                selection = ExcelFormulaSelection(
                    "Book1.xlsx", "Sheet1", "A1", row_count, column_count, ()
                )
                with self.assertRaisesRegex(ValueError, "1 이상"):
                    formula_selection_to_xml(selection)

    def test_rejects_non_integer_dimensions(self):
        for row_count, column_count in ((1.0, 1), (1, "1"), (True, 1)):
            with self.subTest(row_count=row_count, column_count=column_count):
                selection = ExcelFormulaSelection(
                    "Book1.xlsx", "Sheet1", "A1", row_count, column_count, ()
                )
                with self.assertRaisesRegex(ValueError, "정수"):
                    formula_selection_to_xml(selection)

    def test_rejects_full_value_only_selection(self):
        selection = ExcelFormulaSelection(
            "Book1.xlsx",
            "Sheet1",
            "A1:B1",
            1,
            2,
            (ExcelSelectionCell("A1", "1"), ExcelSelectionCell("B1", None)),
        )

        with self.assertRaisesRegex(ValueError, "수식 셀이 없습니다"):
            formula_selection_to_xml(selection)

    def test_rejects_formula_cell_when_current_result_is_missing(self):
        with self.assertRaisesRegex(ValueError, "현재 결과가 없습니다"):
            formula_selection_to_xml(
                self._one_cell(value=None, formula_a1="=SUM(A1:A2)")
            )

    def test_rejects_xml_forbidden_characters_in_every_cell_payload(self):
        cases = (
            ExcelSelectionCell("A\x01", "결과", "=1"),
            ExcelSelectionCell("A1", "결\x01과", "=1"),
            ExcelSelectionCell("A1", "결과", '=CHAR(1)&"\x01"'),
            ExcelSelectionCell("A1", "결과", "=1", "=R\x01C"),
        )
        for cell in cases:
            with self.subTest(cell=cell):
                selection = ExcelFormulaSelection(
                    "Book1.xlsx", "Sheet1", "A1", 1, 1, (cell,)
                )
                with self.assertRaisesRegex(ValueError, "XML 1.0"):
                    formula_selection_to_xml(selection)

    def test_rejects_formula_text_over_the_total_limit(self):
        selection = self._one_cell(value="result", formula_a1="=1234")

        with patch(
            "tablemark.converter.formula_export.MAX_FORMULA_CHARACTERS", 4
        ), self.assertRaisesRegex(ValueError, "수식 내용이 너무 큽니다"):
            formula_selection_to_xml(selection)

    def test_rejects_value_text_over_the_total_limit(self):
        selection = self._one_cell(value="12345", formula_a1="=1")

        with patch(
            "tablemark.converter.formula_export.MAX_VALUE_CHARACTERS", 4
        ), self.assertRaisesRegex(ValueError, "표 값 내용이 너무 큽니다"):
            formula_selection_to_xml(selection)

    def test_rejects_serialized_xml_over_the_byte_limit(self):
        selection = self._one_cell(value="result", formula_a1="=1")

        with patch(
            "tablemark.converter.formula_export.MAX_XML_BYTES", 10
        ), self.assertRaisesRegex(ValueError, "표 XML이 너무 큽니다"):
            formula_selection_to_xml(selection)

    @staticmethod
    def _one_cell(
        *,
        value: str | None,
        formula_a1: str | None,
        formula_r1c1: str | None = None,
    ) -> ExcelFormulaSelection:
        return ExcelFormulaSelection(
            workbook="Book1.xlsx",
            sheet="Sheet1",
            address="A1",
            row_count=1,
            column_count=1,
            cells=(ExcelSelectionCell("A1", value, formula_a1, formula_r1c1),),
        )


if __name__ == "__main__":
    unittest.main()
