"""Serialize a selected Excel table, including values and formulas, as XML."""
from __future__ import annotations

from dataclasses import dataclass
import re
import xml.etree.ElementTree as ET


MAX_FORMULA_CHARACTERS = 1_000_000
MAX_VALUE_CHARACTERS = 5_000_000
MAX_XML_BYTES = 10_000_000
MAX_REFERENCE_CELLS = 10_000
MAX_REFERENCE_RANGES = 256
# Two equal reads validate a snapshot; a third is the one bounded retry when
# the first pair differs or Excel reports a transient selection change.
MAX_SNAPSHOT_READS = 3


@dataclass(frozen=True)
class ExcelReferenceCell:
    """One current value read from a formula's direct static A1 reference."""

    address: str
    value: str | None


@dataclass(frozen=True)
class ExcelFormulaReference:
    """One same-workbook A1 range explicitly referenced by a formula."""

    sheet: str
    address: str
    cells: tuple[ExcelReferenceCell, ...]


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
    references: tuple[ExcelFormulaReference, ...] = ()
    references_complete: bool = True


@dataclass(frozen=True)
class ExcelFormulaSelection:
    """Every cell in one rectangular Excel worksheet selection, row-major."""

    workbook: str
    sheet: str
    address: str
    row_count: int
    column_count: int
    cells: tuple[ExcelSelectionCell, ...]


@dataclass(frozen=True)
class FormulaReferenceTarget:
    """A normalized same-workbook range found in one A1-style formula."""

    sheet: str
    address: str
    row_count: int
    column_count: int


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

    reference_ranges = [
        reference for cell in selection.cells for reference in cell.references
    ]
    if len(reference_ranges) > MAX_REFERENCE_RANGES:
        raise ValueError("수식 참조 범위가 너무 많습니다")
    reference_cells = [
        reference_cell
        for reference in reference_ranges
        for reference_cell in reference.cells
    ]
    if len(reference_cells) > MAX_REFERENCE_CELLS:
        raise ValueError("수식 참조 셀이 너무 많습니다")

    value_characters = sum(
        len(cell.value) for cell in selection.cells if cell.value is not None
    ) + sum(
        len(reference_cell.value)
        for reference_cell in reference_cells
        if reference_cell.value is not None
    )
    if value_characters > MAX_VALUE_CHARACTERS:
        raise ValueError("표 값 내용이 너무 큽니다")

    validate_xml_text(selection.workbook)
    validate_xml_text(selection.sheet)
    validate_xml_text(selection.address)

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
            validate_xml_text(cell.address)
            attributes = {"주소": cell.address}
            if cell.value is not None:
                validate_xml_text(cell.value)
                attributes["값"] = cell.value
            if cell.formula_a1 is not None:
                validate_xml_text(cell.formula_a1)
                attributes["수식"] = cell.formula_a1
            if cell.formula_r1c1 is not None:
                validate_xml_text(cell.formula_r1c1)
                attributes["수식R1C1"] = cell.formula_r1c1
            if not cell.references_complete:
                attributes["참조상태"] = "일부"
            cell_element = ET.SubElement(row, "셀", attributes)
            for reference in cell.references:
                validate_xml_text(reference.sheet)
                validate_xml_text(reference.address)
                reference_element = ET.SubElement(
                    cell_element,
                    "참조범위",
                    {"시트": reference.sheet, "주소": reference.address},
                )
                for reference_cell in reference.cells:
                    validate_xml_text(reference_cell.address)
                    reference_attributes = {"주소": reference_cell.address}
                    if reference_cell.value is not None:
                        validate_xml_text(reference_cell.value)
                        reference_attributes["값"] = reference_cell.value
                    ET.SubElement(reference_element, "참조셀", reference_attributes)

    ET.indent(root, space="  ")
    xml = ET.tostring(root, encoding="unicode")
    if len(xml.encode("utf-8")) > MAX_XML_BYTES:
        raise ValueError("표 XML이 너무 큽니다")
    return xml


def validate_xml_text(value: str) -> None:
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


_MAX_EXCEL_ROW = 1_048_576
_MAX_EXCEL_COLUMN = 16_384
_CELL_REFERENCE = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*"
_BARE_SHEET = r"[A-Za-z_\u0080-\uffff][A-Za-z0-9_.\u0080-\uffff]*"
_A1_REFERENCE_PATTERN = re.compile(
    rf"(?<![A-Za-z0-9_.\]\':])"
    rf"(?:(?:'(?P<quoted_sheet>(?:[^']|'')+)'|(?P<bare_sheet>{_BARE_SHEET}))!)?"
    rf"(?P<start>{_CELL_REFERENCE})"
    rf"(?::(?P<end>{_CELL_REFERENCE}))?"
    rf"(?![A-Za-z0-9_.#(])"
)
_POSSIBLE_NAME_PATTERN = re.compile(
    r"[A-Za-z_\u0080-\uffff][A-Za-z0-9_.\u0080-\uffff]*"
)
_DYNAMIC_REFERENCE_PATTERN = re.compile(r"(?i)\b(?:INDIRECT|OFFSET)\s*\(")
_SHEET_REFERENCE = rf"(?:'(?:[^']|'')+'|{_BARE_SHEET})"
_THREE_DIMENSIONAL_PATTERN = re.compile(
    rf"(?i)(?<![A-Za-z0-9_.\]\':])"
    rf"{_SHEET_REFERENCE}\s*:\s*{_SHEET_REFERENCE}!"
    rf"{_CELL_REFERENCE}(?::{_CELL_REFERENCE})?"
    rf"(?![A-Za-z0-9_.#(])"
)


def extract_formula_reference_targets(
    formula_a1: str, current_sheet: str
) -> tuple[tuple[FormulaReferenceTarget, ...], bool]:
    """Extract direct static same-workbook A1 references without evaluating code.

    The boolean reports whether the extraction is complete. Dynamic references,
    structured references, defined names, 3-D references, and external workbook
    references are deliberately left unresolved and make the result partial.
    """

    if not formula_a1.startswith("="):
        return (), False

    masked = list(formula_a1)
    in_string = False
    bracket_depth = 0
    index = 0
    has_brackets = False
    while index < len(masked):
        character = masked[index]
        if character == '"':
            if in_string and index + 1 < len(masked) and masked[index + 1] == '"':
                masked[index] = " "
                masked[index + 1] = " "
                index += 2
                continue
            in_string = not in_string
            masked[index] = " "
        elif in_string:
            masked[index] = " "
        elif character == "[":
            has_brackets = True
            bracket_depth += 1
            masked[index] = " "
        elif bracket_depth:
            masked[index] = " "
            if character == "]":
                bracket_depth -= 1
        index += 1

    searchable = "".join(masked)
    complete = not has_brackets
    if _DYNAMIC_REFERENCE_PATTERN.search(searchable):
        complete = False
    three_dimensional_matches = tuple(
        _THREE_DIMENSIONAL_PATTERN.finditer(searchable)
    )
    if three_dimensional_matches:
        complete = False
        searchable_characters = list(searchable)
        for match in three_dimensional_matches:
            searchable_characters[match.start() : match.end()] = " " * (
                match.end() - match.start()
            )
        searchable = "".join(searchable_characters)

    matched_spans: list[tuple[int, int]] = []
    targets: list[FormulaReferenceTarget] = []
    seen: set[tuple[str, str]] = set()
    for match in _A1_REFERENCE_PATTERN.finditer(searchable):
        original_start = match.start()
        original_qualifier = formula_a1[match.start() : match.start("start")]
        if (
            (original_start > 0 and formula_a1[original_start - 1] == "]")
            or "[" in original_qualifier
            or "]" in original_qualifier
        ):
            complete = False
            continue

        quoted_sheet = match.group("quoted_sheet")
        bare_sheet = match.group("bare_sheet")
        if quoted_sheet is not None:
            sheet = quoted_sheet.replace("''", "'")
        elif bare_sheet is not None:
            sheet = bare_sheet
        else:
            sheet = current_sheet
        if "[" in sheet or "]" in sheet or ":" in sheet:
            complete = False
            continue

        try:
            start_row, start_column = _parse_cell_reference(match.group("start"))
            raw_end = match.group("end")
            if raw_end is None:
                end_row, end_column = start_row, start_column
            else:
                end_row, end_column = _parse_cell_reference(raw_end)
        except ValueError:
            complete = False
            continue

        first_row, last_row = sorted((start_row, end_row))
        first_column, last_column = sorted((start_column, end_column))
        address = _absolute_range_address(
            first_row, first_column, last_row, last_column
        )
        key = (sheet, address)
        if key not in seen:
            seen.add(key)
            targets.append(
                FormulaReferenceTarget(
                    sheet=sheet,
                    address=address,
                    row_count=last_row - first_row + 1,
                    column_count=last_column - first_column + 1,
                )
            )
        matched_spans.append(match.span())

    unmatched = list(searchable)
    for start, end in matched_spans:
        unmatched[start:end] = " " * (end - start)
    unmatched_text = "".join(unmatched)
    for match in _POSSIBLE_NAME_PATTERN.finditer(unmatched_text):
        token = match.group(0)
        following = unmatched_text[match.end() :].lstrip()
        if following.startswith("(") or token.upper() in {"TRUE", "FALSE"}:
            continue
        complete = False
        break

    return tuple(targets), complete


def _parse_cell_reference(value: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)", value)
    if match is None:
        raise ValueError("invalid A1 cell reference")
    column = 0
    for character in match.group(1).upper():
        column = (column * 26) + ord(character) - ord("A") + 1
    row = int(match.group(2))
    if not 1 <= row <= _MAX_EXCEL_ROW or not 1 <= column <= _MAX_EXCEL_COLUMN:
        raise ValueError("A1 cell reference outside Excel bounds")
    return row, column


def _column_name(number: int) -> str:
    characters = []
    while number:
        number, remainder = divmod(number - 1, 26)
        characters.append(chr(ord("A") + remainder))
    return "".join(reversed(characters))


def _absolute_range_address(
    start_row: int, start_column: int, end_row: int, end_column: int
) -> str:
    first = f"${_column_name(start_column)}${start_row}"
    last = f"${_column_name(end_column)}${end_row}"
    return first if first == last else f"{first}:{last}"
