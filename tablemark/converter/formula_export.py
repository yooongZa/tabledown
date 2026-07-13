"""Serialize a selected Excel table, including values and formulas, as XML."""
from __future__ import annotations

from dataclasses import dataclass
import xml.etree.ElementTree as ET


MAX_FORMULA_CHARACTERS = 1_000_000
MAX_VALUE_CHARACTERS = 5_000_000
MAX_XML_BYTES = 10_000_000


@dataclass(frozen=True)
class ExcelSelectionCell:
    """One cell in a selected Excel range.

    ``value`` is the current cell value converted to text.  A truly blank cell
    uses ``None`` so the XML can distinguish it from an empty string.  Formula
    cells may contain both a calculated value and either formula notation.
    """

    address: str
    value: str | None
    formula_a1: str | None = None
    formula_r1c1: str | None = None


@dataclass(frozen=True)
class ExcelFormulaSelection:
    """Every cell in one rectangular Excel worksheet selection, row-major."""

    workbook: str
    sheet: str
    address: str
    row_count: int
    column_count: int
    cells: tuple[ExcelSelectionCell, ...]


def formula_selection_to_xml(selection: ExcelFormulaSelection) -> str:
    """Return a stable, indented XML representation of ``selection``.

    Values, formulas, and control whitespace are encoded as XML attributes so
    no cell content can create an output line beginning with ``=``.  Blank
    cells omit the ``값`` attribute.  The cell tuple must contain the complete
    rectangular selection in row-major order, and the range must include at
    least one formula.

    Raises:
        ValueError: If the selection shape/content is invalid or too large.
    """
    if type(selection.row_count) is not int or type(selection.column_count) is not int:
        raise ValueError("행수와 열수는 정수여야 합니다")
    if selection.row_count <= 0 or selection.column_count <= 0:
        raise ValueError("행수와 열수는 1 이상이어야 합니다")

    expected_cell_count = selection.row_count * selection.column_count
    if len(selection.cells) != expected_cell_count:
        raise ValueError("선택 범위의 행수·열수와 셀 수가 일치하지 않습니다")

    if not any(
        cell.formula_a1 is not None or cell.formula_r1c1 is not None
        for cell in selection.cells
    ):
        raise ValueError("수식 셀이 없습니다")

    formula_characters = sum(
        len(cell.formula_a1 or "") + len(cell.formula_r1c1 or "")
        for cell in selection.cells
    )
    if formula_characters > MAX_FORMULA_CHARACTERS:
        raise ValueError("수식 내용이 너무 큽니다")

    value_characters = sum(
        len(cell.value) for cell in selection.cells if cell.value is not None
    )
    if value_characters > MAX_VALUE_CHARACTERS:
        raise ValueError("표 값 내용이 너무 큽니다")

    _validate_xml_text(selection.workbook)
    _validate_xml_text(selection.sheet)
    _validate_xml_text(selection.address)

    root = ET.Element(
        "표범위",
        {
            "통합문서": selection.workbook,
            "시트": selection.sheet,
            "주소": selection.address,
            "행수": str(selection.row_count),
            "열수": str(selection.column_count),
        },
    )

    for row_index in range(selection.row_count):
        row = ET.SubElement(root, "행", {"인덱스": str(row_index + 1)})
        start = row_index * selection.column_count
        for cell in selection.cells[start : start + selection.column_count]:
            _validate_xml_text(cell.address)
            attributes = {"주소": cell.address}
            if cell.value is not None:
                _validate_xml_text(cell.value)
                attributes["값"] = cell.value
            if cell.formula_a1 is not None:
                _validate_xml_text(cell.formula_a1)
                attributes["수식"] = cell.formula_a1
            if cell.formula_r1c1 is not None:
                _validate_xml_text(cell.formula_r1c1)
                attributes["수식R1C1"] = cell.formula_r1c1
            ET.SubElement(row, "셀", attributes)

    ET.indent(root, space="  ")
    xml = ET.tostring(root, encoding="unicode")
    if len(xml.encode("utf-8")) > MAX_XML_BYTES:
        raise ValueError("표 XML이 너무 큽니다")
    return xml


def _validate_xml_text(value: str) -> None:
    """Reject characters XML 1.0 cannot represent instead of emitting bad XML."""

    for character in value:
        codepoint = ord(character)
        if character in "\t\n\r":
            continue
        if 0x20 <= codepoint <= 0xD7FF:
            continue
        if 0xE000 <= codepoint <= 0xFFFD:
            continue
        if 0x10000 <= codepoint <= 0x10FFFF:
            continue
        raise ValueError("XML 1.0에서 허용되지 않는 문자가 있습니다")
