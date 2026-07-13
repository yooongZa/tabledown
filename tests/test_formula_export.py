from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET
from unittest.mock import patch

from tablemark.converter.formula_export import (
    ExcelFormulaSelection,
    ExcelSelectionCell,
    formula_selection_to_xml,
)


class FormulaExportTests(unittest.TestCase):
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
            '<표범위 통합문서="매출.xlsx" 시트="요약" 주소="A1:C2" 행수="2" 열수="3">\n'
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
