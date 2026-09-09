from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
import hashlib
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET

from tablemark.converter import formula_export as exporter
from tablemark.converter.formula_export import (
    ExcelFormulaReference,
    ExcelFormulaSelection,
    ExcelReferenceCell,
    ExcelSelectionCell,
    FormulaXmlTooLargeError,
    formula_selection_to_ai_xml,
    formula_selection_to_xml,
)


def _simple_selection() -> ExcelFormulaSelection:
    cells = []
    for column, label in zip("ABCD", ("품목", "단가", "수량", "금액")):
        cells.append(ExcelSelectionCell(f"${column}$1", label, value_kind="text"))
    for row, label, price, quantity in ((2, "사과 🍎", "100", "3"), (3, "배", "200", "2")):
        cells.extend((
            ExcelSelectionCell(f"$A${row}", label, value_kind="text"),
            ExcelSelectionCell(f"$B${row}", price, value_kind="number"),
            ExcelSelectionCell(f"$C${row}", quantity, value_kind="number"),
            ExcelSelectionCell(
                f"$D${row}", str(int(price) * int(quantity)),
                f"=B{row}*C{row}", "=RC[-2]*RC[-1]",
                references=(
                    ExcelFormulaReference("매출", f"$B${row}", (
                        ExcelReferenceCell(f"$B${row}", price, "number"),
                    )),
                    ExcelFormulaReference("매출", f"$C${row}", (
                        ExcelReferenceCell(f"$C${row}", quantity, "number"),
                    )),
                ),
                value_kind="number",
            ),
        ))
    return ExcelFormulaSelection(
        "업무.xlsx", "매출", "$A$1:$D$3", 3, 4, tuple(cells),
        calculation_mode="automatic", calculation_state="done",
    )


def _mixed_selection() -> ExcelFormulaSelection:
    same_sheet = ExcelFormulaReference("Sheet1", "$A$1:$B$2", (
        ExcelReferenceCell("$A$1", "1", "number"),
        ExcelReferenceCell("$B$1", "true", "boolean"),
        ExcelReferenceCell("$A$2", None, "blank"),
        ExcelReferenceCell("$B$2", '#N/A', "error"),
    ))
    other_sheet = ExcelFormulaReference("단가 표", "$A$1", (
        ExcelReferenceCell("$A$1", '한글 & <값> "인용"\n🍎', "text"),
    ))
    overlap = ExcelFormulaReference("Sheet1", "$A$1", (
        same_sheet.cells[0],
    ))
    return ExcelFormulaSelection(
        "Book.xlsx", "Sheet1", "$C$3:$F$3", 1, 4,
        (
            ExcelSelectionCell(
                "$C$3", "1", "=SUM(A1:B2)&'단가 표'!A1", "=SUM(R[-2]C[-2]:R[-1]C[-1])",
                references=(same_sheet, other_sheet), value_kind="number",
            ),
            ExcelSelectionCell(
                "$D$3", "", "='단가 표'!A1&A1&SUM(A1:B2)",
                references=(other_sheet, overlap, same_sheet), value_kind="text",
            ),
            ExcelSelectionCell(
                "$E$3", "#N/A", '=A1+INDIRECT("A2")',
                references=(overlap,), references_complete=False, value_kind="error",
                reference_issues=("dynamic_reference", "range_read_failed", "dynamic_reference"),
            ),
            ExcelSelectionCell("$F$3", None, value_kind="blank"),
        ),
        calculation_mode="manual", calculation_state="pending",
    )


class FormulaAiExportTests(unittest.TestCase):
    def assert_preserves_core(self, selection: ExcelFormulaSelection):
        original = ET.fromstring(formula_selection_to_xml(selection))
        compact = ET.fromstring(formula_selection_to_ai_xml(selection))
        original_attributes = dict(original.attrib)
        compact_attributes = dict(compact.attrib)
        self.assertEqual(compact_attributes.pop("형식"), "AI간결수식")
        self.assertEqual(compact_attributes.pop("형식버전"), "1")
        original_attributes.pop("형식버전")
        self.assertEqual(compact_attributes, original_attributes)
        shared = {}
        for reference in compact.findall("참조목록/참조범위"):
            attrs = dict(reference.attrib)
            reference_id = attrs.pop("id")
            self.assertNotIn(reference_id, shared)
            shared[reference_id] = (attrs, [item.attrib for item in reference])
        original_rows = original.findall("행")
        compact_rows = compact.findall("행")
        self.assertEqual(len(original_rows), len(compact_rows))
        for old_row, new_row in zip(original_rows, compact_rows):
            self.assertEqual(old_row.attrib, new_row.attrib)
            self.assertEqual(len(old_row), len(new_row))
            for old_cell, new_cell in zip(old_row, new_row):
                self.assertEqual(old_cell.attrib, new_cell.attrib)
                old_references = [
                    (reference.attrib, [item.attrib for item in reference])
                    for reference in old_cell.findall("참조범위")
                ]
                new_references = [shared[reference.attrib["ref"]] for reference in new_cell]
                self.assertEqual(old_references, new_references)
        return compact

    def test_all_cell_attributes_and_ordered_reference_values_are_recoverable(self):
        for selection in (_simple_selection(), _mixed_selection()):
            with self.subTest(sheet=selection.sheet):
                self.assert_preserves_core(selection)

    def test_exact_targets_share_once_but_overlap_and_sheet_identity_remain(self):
        root = self.assert_preserves_core(_mixed_selection())
        self.assertEqual(
            [(r.attrib["시트"], r.attrib["주소"]) for r in root.findall("참조목록/참조범위")],
            [("Sheet1", "$A$1:$B$2"), ("단가 표", "$A$1"), ("Sheet1", "$A$1")],
        )
        owners = root.findall("행/셀")
        self.assertEqual([r.attrib["ref"] for r in owners[0]], ["r1", "r2"])
        self.assertEqual([r.attrib["ref"] for r in owners[1]], ["r2", "r3", "r1"])

    def test_simple_context_links_original_labels_without_copying_their_values(self):
        root = self.assert_preserves_core(_simple_selection())
        context = root.find("맥락")
        self.assertEqual(context.attrib["상태"], "추정")
        self.assertEqual(context.attrib["행항목상태"], "추정")
        self.assertEqual(
            [node.attrib for node in context.findall("열제목")],
            [
                {"원본셀": f"${column}$1", "적용범위": f"${column}$2:${column}$3"}
                for column in "ABCD"
            ],
        )
        self.assertEqual(
            [node.attrib for node in context.findall("행항목")],
            [{"원본셀": "$A$2", "적용범위": "$B$2:$D$2"},
             {"원본셀": "$A$3", "적용범위": "$B$3:$D$3"}],
        )
        context_xml = ET.tostring(context, encoding="unicode")
        self.assertNotIn("사과", context_xml)
        self.assertNotIn("단가", context_xml)

    def test_context_uses_actual_selection_offset(self):
        base = _simple_selection()
        # Formula text is preserved as supplied; context only follows cell addresses.
        cells = tuple(replace(cell, address=f"${chr(69 + index % 4)}${10 + index // 4}")
                      for index, cell in enumerate(base.cells))
        root = ET.fromstring(formula_selection_to_ai_xml(replace(
            base, address="$E$10:$H$12", cells=cells,
        )))
        self.assertEqual(root.find("맥락/열제목").attrib,
                         {"원본셀": "$E$10", "적용범위": "$E$11:$E$12"})

    def test_ambiguous_headers_never_get_confirmed_context(self):
        base = _simple_selection()
        for replacement in (
            replace(base.cells[1], value="품목"),
            replace(base.cells[1], value=None, value_kind="blank"),
            replace(base.cells[1], value="   "),
            replace(base.cells[1], value="100", value_kind="number"),
            replace(base.cells[1], value="100"),
            replace(base.cells[1], value_kind=None),
            replace(base.cells[1], formula_a1='="단가"'),
            replace(base.cells[1], value="가격\n원"),
        ):
            with self.subTest(cell=replacement):
                cells = base.cells[:1] + (replacement,) + base.cells[2:]
                root = self.assert_preserves_core(replace(base, cells=cells))
                self.assertEqual(root.find("맥락").attrib["상태"], "미확인")
                self.assertEqual(len(root.find("맥락")), 0)

    def test_second_header_band_is_unknown(self):
        base = _simple_selection()
        cells = list(base.cells)
        cells[5] = replace(cells[5], value="소매", value_kind="text")
        root = self.assert_preserves_core(replace(base, cells=tuple(cells)))
        self.assertEqual(root.find("맥락").attrib["사유"], "ambiguous_header_depth")

    def test_full_width_title_and_unselected_labels_are_not_inferred(self):
        base = _simple_selection()
        cells = list(base.cells)
        cells[0] = replace(cells[0], value="매출 현황")
        for index in (1, 2, 3):
            cells[index] = replace(cells[index], value=None, value_kind="blank")
        root = self.assert_preserves_core(replace(base, cells=tuple(cells)))
        self.assertEqual(root.find("맥락").attrib["상태"], "미확인")
        self.assertEqual(len(root.find("맥락")), 0)
        # Selecting only the numeric body has no in-snapshot header evidence.
        root = self.assert_preserves_core(replace(
            base, address="$A$2:$D$3", row_count=2, cells=base.cells[4:],
        ))
        self.assertEqual(root.find("맥락").attrib["상태"], "미확인")

    def test_duplicate_row_labels_do_not_invent_row_context(self):
        base = _simple_selection()
        cells = list(base.cells)
        cells[8] = replace(cells[8], value=cells[4].value)
        root = self.assert_preserves_core(replace(base, cells=tuple(cells)))
        self.assertEqual(root.find("맥락").attrib["열제목상태"], "추정")
        self.assertEqual(root.find("맥락").attrib["행항목상태"], "미확인")
        self.assertEqual(root.findall("맥락/행항목"), [])

    def test_nonrectangular_addresses_do_not_invent_ranges(self):
        base = _simple_selection()
        for address in ("invalid", "$Z$1", "$XFD$1048576"):
            with self.subTest(address=address):
                cells = (replace(base.cells[0], address=address),) + base.cells[1:]
                root = self.assert_preserves_core(replace(base, cells=cells))
                self.assertEqual(root.find("맥락").attrib["사유"], "ambiguous_cell_addresses")
        root = self.assert_preserves_core(replace(base, address="$B$1:$E$3"))
        self.assertEqual(root.find("맥락").attrib["사유"], "ambiguous_cell_addresses")

    def test_single_cell_and_r1c1_only_retains_unavailable_context_and_core(self):
        for formula_a1, formula_r1c1 in (("=1+1", None), (None, "=1+1")):
            selection = ExcelFormulaSelection("B", "S", "$A$1", 1, 1, (
                ExcelSelectionCell("$A$1", "2", formula_a1, formula_r1c1, value_kind="number"),
            ))
            root = self.assert_preserves_core(selection)
            self.assertEqual(root.find("맥락").attrib["상태"], "미확인")
            self.assertEqual(len(root.find("참조목록")), 0)

    def test_conflicting_shared_snapshot_is_rejected(self):
        base = _mixed_selection()
        conflict = replace(base.cells[0].references[0], cells=(ExcelReferenceCell("$A$1", "999"),))
        cells = base.cells[:1] + (replace(base.cells[1], references=(conflict,)),) + base.cells[2:]
        with self.assertRaisesRegex(ValueError, "같은 수식 참조"):
            formula_selection_to_ai_xml(replace(base, cells=cells))

    def test_empty_reference_and_duplicate_owner_links_are_preserved(self):
        reference = ExcelFormulaReference("Data", "$A$1", ())
        selection = ExcelFormulaSelection("B", "S", "$B$1", 1, 1, (
            ExcelSelectionCell("$B$1", "1", "=Data!A1+Data!A1", references=(reference, reference)),
        ))
        root = self.assert_preserves_core(selection)
        self.assertEqual(len(root.find("참조목록")), 1)
        self.assertEqual(len(root.find("행/셀")), 2)

    def test_output_is_deterministic_without_mutating_snapshot(self):
        selection = _mixed_selection()
        before = deepcopy(selection)
        first = formula_selection_to_ai_xml(selection)
        self.assertEqual(first, formula_selection_to_ai_xml(selection))
        self.assertEqual(selection, before)

    def test_substitution_character_budget_preserves_core_and_warning(self):
        with patch.object(exporter, "MAX_SUBSTITUTED_FORMULA_CHARACTERS", 3):
            root = self.assert_preserves_core(_simple_selection())
        formulas = [cell for cell in root.findall("행/셀") if "수식" in cell.attrib]
        self.assertTrue(formulas)
        for cell in formulas:
            self.assertEqual(cell.attrib["값대입수식상태"], "omitted_character_limit")
            self.assertNotIn("값대입수식", cell.attrib)

    def test_byte_limit_retries_only_without_derived_substitutions(self):
        reference = ExcelFormulaReference("S", "$A$1", (
            ExcelReferenceCell("$A$1", "x" * 1000, "text"),
        ))
        selection = ExcelFormulaSelection("B", "S", "$B$1", 1, 1, (
            ExcelSelectionCell("$B$1", "x", "=A1", references=(reference,), value_kind="text"),
        ), calculation_mode="manual", calculation_state="pending")
        with patch.object(exporter, "MAX_SUBSTITUTED_FORMULA_CHARACTERS", 0):
            core_size = len(formula_selection_to_ai_xml(selection).encode("utf-8"))
        with patch.object(exporter, "MAX_XML_BYTES", core_size + 100):
            fallback_xml = formula_selection_to_ai_xml(selection)
            root = ET.fromstring(fallback_xml)
        self.assertEqual(root.attrib["값대입수식상태"], "omitted_xml_size_limit")
        self.assertEqual(root.attrib["계산결과상태"], "calculation_incomplete")
        self.assertEqual(root.find("참조목록/참조범위/참조셀").attrib["값"], "x" * 1000)
        self.assertNotIn("값대입수식", root.find("행/셀").attrib)
        with patch.object(exporter, "MAX_XML_BYTES", len(fallback_xml.encode("utf-8")) - 1):
            with self.assertRaises(FormulaXmlTooLargeError):
                formula_selection_to_ai_xml(selection)

    def test_serialization_stops_before_materializing_all_reference_cells(self):
        reference = ExcelFormulaReference("S", "$A$1:$A$2048", tuple(
            ExcelReferenceCell(f"$A${row}", "한글" * 100, "text")
            for row in range(1, 2049)
        ))
        selection = ExcelFormulaSelection("B", "S", "$B$1", 1, 1, (
            ExcelSelectionCell("$B$1", "x", "=SUM(A1:A2048)", references=(reference,)),
        ))
        with patch.object(exporter, "MAX_XML_BYTES", 2000), patch.object(
            exporter, "_serialized_empty_element", wraps=exporter._serialized_empty_element,
        ) as serializer:
            with self.assertRaises(FormulaXmlTooLargeError):
                formula_selection_to_ai_xml(selection)
        self.assertLess(serializer.call_count, 30)

    def test_shared_reference_heavy_input_uses_less_utf8_bytes(self):
        reference = ExcelFormulaReference("Data", "$A$1:$A$128", tuple(
            ExcelReferenceCell(f"$A${row}", str(row), "number") for row in range(1, 129)
        ))
        selection = ExcelFormulaSelection("B", "S", "$B$1:$B$30", 30, 1, tuple(
            ExcelSelectionCell(f"$B${row}", "8256", "=SUM(Data!A1:A128)",
                               references=(reference,), value_kind="number")
            for row in range(1, 31)
        ))
        self.assert_preserves_core(selection)
        old_size = len(formula_selection_to_xml(selection).encode("utf-8"))
        new_size = len(formula_selection_to_ai_xml(selection).encode("utf-8"))
        self.assertLess(new_size, old_size / 3)

    def test_original_xml_bytes_remain_unchanged(self):
        # Recorded from the task baseline before the compact serializer existed.
        self.assertEqual(
            hashlib.sha256(formula_selection_to_xml(_mixed_selection()).encode("utf-8")).hexdigest(),
            "9bd3ee215e51b9df1b3674bbaba74417abc545091f5183db98bf94301010983d",
        )

    def test_existing_validation_limits_and_invalid_content_apply_to_both_formats(self):
        base = _mixed_selection()
        cases = (
            replace(base, row_count=True),
            replace(base, row_count=0),
            replace(base, cells=base.cells[:1]),
            replace(base, workbook="invalid\x00"),
            replace(base, cells=(replace(base.cells[0], value_kind="invalid"),) + base.cells[1:]),
            replace(base, cells=(replace(base.cells[0], value=None),) + base.cells[1:]),
        )
        for selection in cases:
            for writer in (formula_selection_to_xml, formula_selection_to_ai_xml):
                with self.subTest(writer=writer.__name__, selection=selection):
                    with self.assertRaises(ValueError):
                        writer(selection)
        for constant in ("MAX_FORMULA_CHARACTERS", "MAX_VALUE_CHARACTERS", "MAX_REFERENCE_RANGES", "MAX_REFERENCE_CELLS"):
            with self.subTest(limit=constant), patch.object(exporter, constant, 0):
                with self.assertRaises(ValueError):
                    formula_selection_to_ai_xml(base)


if __name__ == "__main__":
    unittest.main()
