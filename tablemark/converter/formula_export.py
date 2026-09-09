"""Serialize a selected Excel table, including values and formulas, as XML."""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from xml.sax.saxutils import escape


MAX_FORMULA_CHARACTERS = 1_000_000
MAX_VALUE_CHARACTERS = 5_000_000
MAX_XML_BYTES = 10_000_000
MAX_REFERENCE_CELLS = 10_000
MAX_REFERENCE_RANGES = 256
# A substituted formula is only a readability aid.  Keep its aggregate size
# bounded so duplicating a large referenced range never turns a valid export
# into an unexpectedly large in-memory string.
MAX_SUBSTITUTED_FORMULA_CHARACTERS = 1_000_000
# Two equal reads validate a snapshot; a third is the one bounded retry when
# the first pair differs or Excel reports a transient selection change.
MAX_SNAPSHOT_READS = 3

_VALUE_KIND_TO_XML = {
    "blank": "blank",
    "number": "number",
    "text": "text",
    "boolean": "boolean",
    "error": "error",
}


class FormulaXmlTooLargeError(ValueError):
    """The core formula XML still exceeded the established 10 MB limit."""


class _SubstitutionCharacterLimit(RuntimeError):
    """A valid derived substitution exceeded its dedicated character budget."""


class _BoundedFormulaXmlLines:
    """Accumulate XML lines without ever materializing an oversized document."""

    def __init__(self, max_bytes: int) -> None:
        self._max_bytes = max_bytes
        self._byte_count = 0
        self._lines: list[str] = []

    def append(self, line: str) -> None:
        separator_bytes = 1 if self._lines else 0
        next_byte_count = (
            self._byte_count + separator_bytes + len(line.encode("utf-8"))
        )
        if next_byte_count > self._max_bytes:
            raise FormulaXmlTooLargeError("표 XML이 너무 큽니다")
        self._lines.append(line)
        self._byte_count = next_byte_count

    def ensure_minimum_line_characters(self, minimum_characters: int) -> None:
        """Reject a line before escaping a value that cannot possibly fit.

        Every Unicode code point takes at least one UTF-8 byte and XML escaping
        only grows text, so the character count is a safe lower bound.
        """

        separator_bytes = 1 if self._lines else 0
        if (
            self._byte_count + separator_bytes + minimum_characters
            > self._max_bytes
        ):
            raise FormulaXmlTooLargeError("표 XML이 너무 큽니다")

    def render(self) -> str:
        return "\n".join(self._lines)


@dataclass(frozen=True)
class ExcelReferenceCell:
    """One current value read from a formula's direct static A1 reference."""

    address: str
    value: str | None
    value_kind: str | None = None


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
    value_kind: str | None = None
    reference_issues: tuple[str, ...] = ()


@dataclass(frozen=True)
class ExcelFormulaSelection:
    """Every cell in one rectangular Excel worksheet selection, row-major."""

    workbook: str
    sheet: str
    address: str
    row_count: int
    column_count: int
    cells: tuple[ExcelSelectionCell, ...]
    calculation_mode: str | None = None
    calculation_state: str | None = None


@dataclass(frozen=True)
class FormulaReferenceTarget:
    """A normalized same-workbook range found in one A1-style formula."""

    sheet: str
    address: str
    row_count: int
    column_count: int


@dataclass(frozen=True)
class _FormulaReferenceOccurrence:
    """One source occurrence and its normalized same-workbook target."""

    start: int
    end: int
    target: FormulaReferenceTarget


def formula_copy_notice_key(selection: ExcelFormulaSelection) -> str | None:
    """Describe actionable limits of a copied snapshot for both platforms."""

    partial_references = any(
        not cell.references_complete for cell in selection.cells
    )
    calculation_incomplete = selection.calculation_state in {
        "calculating", "pending"
    }
    if partial_references and calculation_incomplete:
        return "partial_references_and_calculation"
    if partial_references:
        return "partial_references"
    if calculation_incomplete:
        return "calculation_incomplete"
    return None


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
    _validate_formula_selection(selection)

    try:
        return _render_formula_xml(
            selection,
            include_substitutions=True,
            substitution_status=None,
        )
    except FormulaXmlTooLargeError:
        # 값대입수식은 설명용 파생 정보다. 원본 값·수식·타입·계산
        # 상태·참조 누락 사유는 조용히 제거하지 않고 그대로 재시도한다.
        pass
    return _render_formula_xml(
        selection,
        include_substitutions=False,
        substitution_status="omitted_xml_size_limit",
    )


def formula_selection_to_ai_xml(selection: ExcelFormulaSelection) -> str:
    """Return AI-oriented XML without repeating exact shared reference ranges.

    All selected cells and core formula attributes are retained. Context only
    links addresses within the existing snapshot and is explicitly inferred;
    no additional worksheet reads or recalculation are performed.
    """
    references = _validate_formula_selection(selection)
    try:
        return _render_formula_xml(
            selection,
            include_substitutions=True,
            substitution_status=None,
            shared_references=references,
        )
    except FormulaXmlTooLargeError:
        pass
    return _render_formula_xml(
        selection,
        include_substitutions=False,
        substitution_status="omitted_xml_size_limit",
        shared_references=references,
    )


def _validate_formula_selection(
    selection: ExcelFormulaSelection,
) -> dict[tuple[str, str], ExcelFormulaReference]:
    """Validate either output format against the same snapshot contract."""
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

    if any(
        (cell.formula_a1 is not None or cell.formula_r1c1 is not None)
        and cell.value is None
        for cell in selection.cells
    ):
        # Formula errors and formulas returning "" are represented by a text
        # value.  ``None`` here means the reader silently lost the current
        # result, so emitting a formula without its result would be misleading.
        raise ValueError("수식 셀의 현재 결과가 없습니다")

    formula_characters = sum(
        len(cell.formula_a1 or "") + len(cell.formula_r1c1 or "")
        for cell in selection.cells
    )
    if formula_characters > MAX_FORMULA_CHARACTERS:
        raise ValueError("수식 내용이 너무 큽니다")

    unique_references: dict[
        tuple[str, str], ExcelFormulaReference
    ] = {}
    for cell in selection.cells:
        for reference in cell.references:
            key = (reference.sheet, reference.address)
            existing = unique_references.get(key)
            if existing is not None and existing != reference:
                raise ValueError("같은 수식 참조 범위의 값이 일치하지 않습니다")
            unique_references[key] = reference
    if len(unique_references) > MAX_REFERENCE_RANGES:
        raise ValueError("수식 참조 범위가 너무 많습니다")
    reference_cell_count = sum(
        len(reference.cells) for reference in unique_references.values()
    )
    if reference_cell_count > MAX_REFERENCE_CELLS:
        raise ValueError("수식 참조 셀이 너무 많습니다")

    value_characters = sum(
        len(cell.value) for cell in selection.cells if cell.value is not None
    ) + sum(
        len(reference_cell.value)
        for reference in unique_references.values()
        for reference_cell in reference.cells
        if reference_cell.value is not None
    )
    if value_characters > MAX_VALUE_CHARACTERS:
        raise ValueError("표 값 내용이 너무 큽니다")

    validate_xml_text(selection.workbook)
    validate_xml_text(selection.sheet)
    validate_xml_text(selection.address)
    for state in (selection.calculation_mode, selection.calculation_state):
        if state is not None:
            validate_xml_text(state)
    for cell in selection.cells:
        validate_xml_text(cell.address)
        if cell.value is not None:
            validate_xml_text(cell.value)
        if cell.formula_a1 is not None:
            validate_xml_text(cell.formula_a1)
        if cell.formula_r1c1 is not None:
            validate_xml_text(cell.formula_r1c1)
        if cell.value_kind is not None and cell.value_kind not in _VALUE_KIND_TO_XML:
            raise ValueError("셀 값 종류가 올바르지 않습니다")
        for issue in cell.reference_issues:
            if not isinstance(issue, str):
                raise ValueError("수식 참조 누락 사유가 올바르지 않습니다")
            validate_xml_text(issue)
    for reference in unique_references.values():
        validate_xml_text(reference.sheet)
        validate_xml_text(reference.address)
        for reference_cell in reference.cells:
            validate_xml_text(reference_cell.address)
            if reference_cell.value is not None:
                validate_xml_text(reference_cell.value)
            if (
                reference_cell.value_kind is not None
                and reference_cell.value_kind not in _VALUE_KIND_TO_XML
            ):
                raise ValueError("참조 셀 값 종류가 올바르지 않습니다")

    return unique_references


def _render_formula_xml(
    selection: ExcelFormulaSelection,
    *,
    include_substitutions: bool,
    substitution_status: str | None,
    shared_references: dict[tuple[str, str], ExcelFormulaReference] | None = None,
) -> str:
    """Render one already validated selection under the hard byte cap."""

    lines = _BoundedFormulaXmlLines(MAX_XML_BYTES)
    root_attributes = {
        "통합문서": selection.workbook,
        "시트": selection.sheet,
        "주소": selection.address,
        "행수": str(selection.row_count),
        "열수": str(selection.column_count),
        "형식버전": "2",
        "계산모드": selection.calculation_mode or "unknown",
        "계산상태": selection.calculation_state or "unavailable",
        "계산결과상태": _calculation_result_status(selection),
        "값기준": "Excel현재원시값",
        "표시정보상태": "미포함",
        "병합정보상태": "미포함",
    }
    if shared_references is not None:
        root_attributes["형식버전"] = "1"
        root_attributes["형식"] = "AI간결수식"
    if substitution_status is not None:
        root_attributes["값대입수식상태"] = substitution_status
    _append_open_element(lines, "", "표범위", root_attributes)

    reference_ids: dict[tuple[str, str], str] = {}
    if shared_references is not None:
        _append_formula_context(lines, selection)
        _append_open_element(lines, "  ", "참조목록", {})
        for index, (key, reference) in enumerate(shared_references.items(), 1):
            reference_id = f"r{index}"
            reference_ids[key] = reference_id
            _append_formula_reference(lines, "    ", reference, reference_id)
        lines.append("  </참조목록>")

    substituted_formula_characters = 0
    for row_index in range(selection.row_count):
        _append_open_element(
            lines,
            "  ",
            "행",
            {"인덱스": str(row_index + 1)},
        )
        start = row_index * selection.column_count
        for cell in selection.cells[start : start + selection.column_count]:
            attributes = {"주소": cell.address}
            if cell.value is not None:
                attributes["값"] = cell.value
            if cell.value_kind is not None:
                attributes["값종류"] = _VALUE_KIND_TO_XML[cell.value_kind]
            if cell.formula_a1 is not None:
                attributes["수식"] = cell.formula_a1
            if cell.formula_r1c1 is not None:
                attributes["수식R1C1"] = cell.formula_r1c1
            if include_substitutions:
                remaining_substitution_characters = (
                    MAX_SUBSTITUTED_FORMULA_CHARACTERS
                    - substituted_formula_characters
                )
                try:
                    substituted_formula = _substitute_formula_reference_values(
                        cell,
                        selection.sheet,
                        max_characters=remaining_substitution_characters,
                    )
                except _SubstitutionCharacterLimit:
                    substituted_formula = None
                    attributes["값대입수식상태"] = (
                        "omitted_character_limit"
                    )
                if substituted_formula is not None:
                    validate_xml_text(substituted_formula)
                    attributes["값대입수식"] = substituted_formula
                    attributes["값대입수식동등성"] = "보장안함"
                    substituted_formula_characters += len(substituted_formula)
            if not cell.references_complete:
                attributes["참조상태"] = "일부"
                issues = cell.reference_issues or ("unspecified",)
                attributes["참조누락이유"] = ",".join(dict.fromkeys(issues))
                attributes["참조포함범위수"] = str(len(cell.references))

            if not cell.references:
                _append_empty_element(lines, "    ", "셀", attributes)
                continue

            _append_open_element(lines, "    ", "셀", attributes)
            for reference in cell.references:
                if shared_references is not None:
                    _append_empty_element(
                        lines,
                        "      ",
                        "참조",
                        {"ref": reference_ids[(reference.sheet, reference.address)]},
                    )
                else:
                    _append_formula_reference(lines, "      ", reference)
            lines.append("    </셀>")
        lines.append("  </행>")
    lines.append("</표범위>")
    return lines.render()


def _append_formula_reference(
    lines: _BoundedFormulaXmlLines,
    indentation: str,
    reference: ExcelFormulaReference,
    reference_id: str | None = None,
) -> None:
    attributes = {"시트": reference.sheet, "주소": reference.address}
    if reference_id is not None:
        attributes["id"] = reference_id
    if not reference.cells:
        _append_empty_element(lines, indentation, "참조범위", attributes)
        return
    _append_open_element(lines, indentation, "참조범위", attributes)
    for cell in reference.cells:
        cell_attributes = {"주소": cell.address}
        if cell.value is not None:
            cell_attributes["값"] = cell.value
        if cell.value_kind is not None:
            cell_attributes["값종류"] = _VALUE_KIND_TO_XML[cell.value_kind]
        _append_empty_element(
            lines, indentation + "  ", "참조셀", cell_attributes
        )
    lines.append(indentation + "</참조범위>")


def _context_label(cell: ExcelSelectionCell) -> str | None:
    """Only short, native text literals can be tentative table labels."""
    if (
        cell.value_kind != "text"
        or cell.formula_a1 is not None
        or cell.formula_r1c1 is not None
        or cell.value is None
        or len(cell.value) > 128
        or any(character in cell.value for character in "\r\n\t")
    ):
        return None
    label = cell.value.strip()
    if not label:
        return None
    try:
        Decimal(label)
    except InvalidOperation:
        return label
    return None


def _append_formula_context(
    lines: _BoundedFormulaXmlLines, selection: ExcelFormulaSelection
) -> None:
    """Conservatively link simple headers/row labels to their original cells.

    Addresses, rather than copied labels, keep this metadata small. Merged
    topology is unavailable in a formula snapshot, so a text-like second row
    or ambiguous first row deliberately yields no inferred column labels.
    """
    attributes = {"상태": "미확인", "기준": "선택범위내단순헤더"}
    if selection.row_count < 2 or selection.column_count < 2:
        attributes["사유"] = "insufficient_rows_or_columns"
        _append_empty_element(lines, "  ", "맥락", attributes)
        return
    try:
        first_row, first_column = _parse_cell_reference(selection.cells[0].address)
        last_row = first_row + selection.row_count - 1
        last_column = first_column + selection.column_count - 1
        if last_row > _MAX_EXCEL_ROW or last_column > _MAX_EXCEL_COLUMN:
            raise ValueError("selection outside worksheet")
        bounds = selection.address.split(":")
        if (
            len(bounds) not in {1, 2}
            or _parse_cell_reference(bounds[0]) != (first_row, first_column)
            or _parse_cell_reference(bounds[-1]) != (last_row, last_column)
        ):
            raise ValueError("selection and cell addresses disagree")
        for index, cell in enumerate(selection.cells):
            row, column = divmod(index, selection.column_count)
            if _parse_cell_reference(cell.address) != (
                first_row + row, first_column + column
            ):
                raise ValueError("nonrectangular cell addresses")
    except ValueError:
        attributes["사유"] = "ambiguous_cell_addresses"
        _append_empty_element(lines, "  ", "맥락", attributes)
        return

    headers = selection.cells[: selection.column_count]
    labels = [_context_label(cell) for cell in headers]
    if None in labels or len(set(labels)) != len(labels):
        attributes["사유"] = "ambiguous_header_labels"
        _append_empty_element(lines, "  ", "맥락", attributes)
        return
    # A second header band often consists of more text. Require an immediate
    # data-like cell beneath every candidate value-column label instead.
    second_row = selection.cells[
        selection.column_count : 2 * selection.column_count
    ]
    if any(
        cell.formula_a1 is None and cell.formula_r1c1 is None
        and cell.value_kind not in {"number", "boolean", "error"}
        for cell in second_row[1:]
    ):
        attributes["사유"] = "ambiguous_header_depth"
        _append_empty_element(lines, "  ", "맥락", attributes)
        return

    row_labels = selection.cells[
        selection.column_count :: selection.column_count
    ]
    row_label_values = [_context_label(cell) for cell in row_labels]
    have_row_labels = (
        None not in row_label_values
        and len(set(row_label_values)) == len(row_label_values)
    )
    attributes["상태"] = "추정"
    attributes["열제목상태"] = "추정"
    attributes["행항목상태"] = "추정" if have_row_labels else "미확인"
    if not have_row_labels:
        attributes["행항목사유"] = "ambiguous_row_labels"
    _append_open_element(lines, "  ", "맥락", attributes)
    for index, cell in enumerate(headers):
        _append_empty_element(
            lines, "    ", "열제목",
            {
                "원본셀": cell.address,
                "적용범위": _absolute_range_address(
                    first_row + 1, first_column + index,
                    last_row, first_column + index,
                ),
            },
        )
    if have_row_labels:
        for index, cell in enumerate(row_labels, 1):
            _append_empty_element(
                lines, "    ", "행항목",
                {
                    "원본셀": cell.address,
                    "적용범위": _absolute_range_address(
                        first_row + index, first_column + 1,
                        first_row + index, last_column,
                    ),
                },
            )
    lines.append("  </맥락>")


def _serialized_empty_element(tag: str, attributes: dict[str, str]) -> str:
    """Serialize one trusted schema tag without ElementTree's large-value copies."""

    parts = [f"<{tag}"]
    attribute_entities = {
        '"': "&quot;",
        "\t": "&#09;",
        "\n": "&#10;",
        "\r": "&#13;",
    }
    for name, value in attributes.items():
        escaped_value = escape(value, attribute_entities)
        parts.extend((" ", name, '="', escaped_value, '"'))
    parts.append(" />")
    return "".join(parts)


def _serialized_open_element(tag: str, attributes: dict[str, str]) -> str:
    serialized = _serialized_empty_element(tag, attributes)
    if not serialized.endswith(" />"):
        raise ValueError("XML 시작 태그를 만들 수 없습니다")
    return serialized[:-3] + ">"


def _append_empty_element(
    lines: _BoundedFormulaXmlLines,
    indentation: str,
    tag: str,
    attributes: dict[str, str],
) -> None:
    lines.ensure_minimum_line_characters(
        _minimum_element_characters(
            indentation, tag, attributes, closing_characters=3
        )
    )
    lines.append(
        indentation + _serialized_empty_element(tag, attributes)
    )


def _append_open_element(
    lines: _BoundedFormulaXmlLines,
    indentation: str,
    tag: str,
    attributes: dict[str, str],
) -> None:
    lines.ensure_minimum_line_characters(
        _minimum_element_characters(
            indentation, tag, attributes, closing_characters=1
        )
    )
    lines.append(
        indentation + _serialized_open_element(tag, attributes)
    )


def _minimum_element_characters(
    indentation: str,
    tag: str,
    attributes: dict[str, str],
    *,
    closing_characters: int,
) -> int:
    # "<tag" + each ' name="value"' + either ">" or " />".
    return (
        len(indentation)
        + 1
        + len(tag)
        + sum(
            4 + len(name) + len(value)
            for name, value in attributes.items()
        )
        + closing_characters
    )


def _calculation_result_status(selection: ExcelFormulaSelection) -> str:
    """Describe what the snapshot proves without claiming recalculation."""

    if selection.calculation_state == "done":
        return "snapshot_stable"
    if selection.calculation_state in {"calculating", "pending"}:
        return "calculation_incomplete"
    return "freshness_unverified"


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


def normalize_excel_number_text(value: str) -> str:
    """Expand one Excel numeric value without changing numeric-looking text.

    Callers must only pass values whose native Excel type was numeric.  That
    distinction is important: a literal text cell containing ``3.15E+5`` must
    remain exactly that text.
    """

    try:
        number = Decimal(value.strip())
    except (InvalidOperation, ValueError):
        return value
    if not number.is_finite():
        return value
    plain = format(number, "f")
    if "." in plain:
        plain = plain.rstrip("0").rstrip(".")
    if plain in {"-0", "+0", ""}:
        return "0"
    return plain


_MAX_EXCEL_ROW = 1_048_576
_MAX_EXCEL_COLUMN = 16_384
_CELL_REFERENCE = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*"
_BARE_SHEET = r"[A-Za-z_\u0080-\uffff][A-Za-z0-9_.\u0080-\uffff]*"
_A1_REFERENCE_PATTERN = re.compile(
    rf"(?<![\w.\\\]\':])"
    rf"(?:(?:'(?P<quoted_sheet>(?:[^']|'')+)'|(?P<bare_sheet>{_BARE_SHEET}))!)?"
    rf"(?P<start>{_CELL_REFERENCE})"
    rf"(?:[ \t\r\n]*:[ \t\r\n]*(?P<end>{_CELL_REFERENCE}))?"
    rf"(?![\w.#(])"
)
_WHOLE_ROW_OR_COLUMN_REFERENCE_PATTERN = re.compile(
    r"(?i)(?<![A-Za-z0-9_.])(?:\$?[1-9][0-9]*\s*:\s*\$?[1-9][0-9]*|"
    r"\$?[A-Z]{1,3}\s*:\s*\$?[A-Z]{1,3})(?![A-Za-z0-9_.])"
)
_POSSIBLE_NAME_PATTERN = re.compile(
    r"[A-Za-z_\u0080-\uffff][A-Za-z0-9_.\u0080-\uffff]*"
)
_NUMERIC_LITERAL_PATTERN = re.compile(
    r"(?i)(?<![A-Za-z0-9_.])(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)"
    r"(?:E[+-]?[0-9]+)?(?![A-Za-z0-9_.])"
)
_ERROR_LITERAL_PATTERN = re.compile(
    # Slash is part of #N/A and #DIV/0!, but an operator after other
    # error literals. Do not swallow the A1 reference in #N/A/A1.
    r"(?i)#(?:N/A|DIV/0!|[A-Z][A-Z0-9_]*(?:[!?])?)"
)
_QUOTED_SHEET_PREFIX_PATTERN = re.compile(r"'(?:[^']|'')+'!")
_BROKEN_A1_REFERENCE_PATTERN = re.compile(
    rf"(?i)#REF!{_CELL_REFERENCE}"
    rf"(?:\s*:\s*{_CELL_REFERENCE})?(?![\w.#(])"
)
_DYNAMIC_REFERENCE_PATTERN = re.compile(r"(?i)\b(?:INDIRECT|OFFSET)\s*\(")
_CALCULATED_RANGE_PATTERN = re.compile(
    r":\s*(?:\(\s*)*@?[\w.]+\s*\(|\)\s*:"
)
_SHEET_REFERENCE = rf"(?:'(?:[^']|'')+'|{_BARE_SHEET})"
_THREE_DIMENSIONAL_PATTERN = re.compile(
    rf"(?i)(?<![A-Za-z0-9_.\]\':])"
    rf"{_SHEET_REFERENCE}\s*:\s*{_SHEET_REFERENCE}!"
    rf"{_CELL_REFERENCE}(?:[ \t\r\n]*:[ \t\r\n]*{_CELL_REFERENCE})?"
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

    targets, issues = analyze_formula_references(formula_a1, current_sheet)
    return targets, not issues


def analyze_formula_references(
    formula_a1: str, current_sheet: str
) -> tuple[tuple[FormulaReferenceTarget, ...], tuple[str, ...]]:
    """Return supported targets plus stable reasons for unresolved syntax."""

    occurrences, issues = _formula_reference_occurrences(
        formula_a1, current_sheet
    )
    targets: list[FormulaReferenceTarget] = []
    seen: set[tuple[str, str]] = set()
    for occurrence in occurrences:
        key = (occurrence.target.sheet, occurrence.target.address)
        if key not in seen:
            seen.add(key)
            targets.append(occurrence.target)
    return tuple(targets), issues


def _formula_reference_occurrences(
    formula_a1: str, current_sheet: str
) -> tuple[tuple[_FormulaReferenceOccurrence, ...], tuple[str, ...]]:
    """Return every supported A1 occurrence with its source span."""

    if not formula_a1.startswith("="):
        return (), ("invalid_formula",)

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
    issues: list[str] = []

    def add_issue(issue: str) -> None:
        if issue not in issues:
            issues.append(issue)

    # Mask error literals before A1 matching: #REF!B2 must never become
    # a read of B2 on a real worksheet named REF. Quoted worksheet names
    # may legitimately contain error-looking text, so exclude those first.
    error_searchable = list(searchable)
    for match in _QUOTED_SHEET_PREFIX_PATTERN.finditer(searchable):
        error_searchable[match.start() : match.end()] = " " * (
            match.end() - match.start()
        )
    error_searchable_text = "".join(error_searchable)
    searchable_characters = list(searchable)
    for match in _BROKEN_A1_REFERENCE_PATTERN.finditer(error_searchable_text):
        add_issue("invalid_a1_reference")
        searchable_characters[match.start() : match.end()] = " " * (
            match.end() - match.start()
        )
    for match in _ERROR_LITERAL_PATTERN.finditer(error_searchable_text):
        searchable_characters[match.start() : match.end()] = " " * (
            match.end() - match.start()
        )
    searchable = "".join(searchable_characters)

    if has_brackets:
        add_issue("external_or_structured_reference")
    if _DYNAMIC_REFERENCE_PATTERN.search(searchable):
        add_issue("dynamic_reference")
    # INDEX/CHOOSE and other functions can return one endpoint of a range.
    # The literal A1 arguments alone do not cover that resulting rectangle.
    # Preserve those known references without claiming complete enrichment.
    if _CALCULATED_RANGE_PATTERN.search(searchable):
        add_issue("calculated_range_reference")
    if _WHOLE_ROW_OR_COLUMN_REFERENCE_PATTERN.search(searchable):
        add_issue("whole_row_or_column_reference")
    three_dimensional_matches = tuple(
        _THREE_DIMENSIONAL_PATTERN.finditer(searchable)
    )
    if three_dimensional_matches:
        add_issue("three_dimensional_reference")
        searchable_characters = list(searchable)
        for match in three_dimensional_matches:
            searchable_characters[match.start() : match.end()] = " " * (
                match.end() - match.start()
            )
        searchable = "".join(searchable_characters)

    matched_spans: list[tuple[int, int]] = []
    occurrences: list[_FormulaReferenceOccurrence] = []
    for match in _A1_REFERENCE_PATTERN.finditer(searchable):
        original_start = match.start()
        original_qualifier = formula_a1[match.start() : match.start("start")]
        if (
            (original_start > 0 and formula_a1[original_start - 1] == "]")
            or "[" in original_qualifier
            or "]" in original_qualifier
        ):
            add_issue("external_or_structured_reference")
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
            add_issue("external_or_structured_reference")
            continue

        try:
            start_row, start_column = _parse_cell_reference(match.group("start"))
            raw_end = match.group("end")
            if raw_end is None:
                end_row, end_column = start_row, start_column
            else:
                end_row, end_column = _parse_cell_reference(raw_end)
        except ValueError:
            add_issue("invalid_a1_reference")
            continue

        first_row, last_row = sorted((start_row, end_row))
        first_column, last_column = sorted((start_column, end_column))
        address = _absolute_range_address(
            first_row, first_column, last_row, last_column
        )
        occurrences.append(
            _FormulaReferenceOccurrence(
                start=match.start(),
                end=match.end(),
                target=FormulaReferenceTarget(
                    sheet=sheet,
                    address=address,
                    row_count=last_row - first_row + 1,
                    column_count=last_column - first_column + 1,
                ),
            )
        )
        matched_spans.append(match.span())
        # A colon inside the matched span belongs to a supported A1 range.
        # A colon just outside it has an unresolved endpoint, including a
        # parenthesized cell or an error literal masked above. Inspect only
        # adjacent whitespace, without copying the remaining formula text.
        before = match.start() - 1
        # After an unresolved colon, matching can start at A in $A$1 because
        # the leading $ itself was preceded by the colon boundary guard.
        if before >= 0 and searchable[before] == "$":
            before -= 1
        while before >= 0 and searchable[before].isspace():
            before -= 1
        after = match.end()
        while after < len(searchable) and searchable[after].isspace():
            after += 1
        if (
            (before >= 0 and searchable[before] == ":")
            or (after < len(searchable) and searchable[after] == ":")
        ):
            add_issue("calculated_range_reference")

    unmatched = list(searchable)
    for start, end in matched_spans:
        unmatched[start:end] = " " * (end - start)
    unmatched_text = "".join(unmatched)
    unmatched_without_literals = list(unmatched_text)
    for match in _NUMERIC_LITERAL_PATTERN.finditer(unmatched_text):
        unmatched_without_literals[match.start() : match.end()] = " " * (
            match.end() - match.start()
        )
    unmatched_text = "".join(unmatched_without_literals)
    for match in _POSSIBLE_NAME_PATTERN.finditer(unmatched_text):
        token = match.group(0)
        following = unmatched_text[match.end() :].lstrip()
        # Excel exposes no stable, locale-independent built-in function list
        # here. Treat callable tokens as functions so newly added and uncommon
        # built-ins do not become false partials. The derived expression still
        # substitutes only A1 references written explicitly in the formula.
        if following.startswith("(") or token.upper() in {"TRUE", "FALSE"}:
            continue
        add_issue("defined_name_or_unsupported_syntax")
        break

    return tuple(occurrences), tuple(issues)


def _substitute_formula_reference_values(
    cell: ExcelSelectionCell,
    current_sheet: str,
    *,
    max_characters: int,
) -> str | None:
    """Build a bounded, explanatory formula with direct references inlined.

    This derived expression is for human/LLM inspection only.  It is not sent
    back to Excel and is not guaranteed to be calculation-equivalent for every
    Excel operator.  The original A1/R1C1 formulas remain authoritative.
    """

    if cell.formula_a1 is None or not cell.references_complete:
        return None
    occurrences, issues = _formula_reference_occurrences(
        cell.formula_a1, current_sheet
    )
    if issues or not occurrences:
        return None

    references = {
        (reference.sheet, reference.address): reference
        for reference in cell.references
    }
    occurrence_keys = {
        (occurrence.target.sheet, occurrence.target.address)
        for occurrence in occurrences
    }
    if (
        len(references) != len(cell.references)
        or set(references) != occurrence_keys
    ):
        return None
    if max_characters <= 0:
        raise _SubstitutionCharacterLimit

    pieces: list[str] = []
    current_length = 0
    source_index = 0
    for occurrence in occurrences:
        key = (occurrence.target.sheet, occurrence.target.address)
        reference = references.get(key)
        if reference is None:
            return None
        prefix = cell.formula_a1[source_index : occurrence.start]
        remaining_characters = max_characters - current_length - len(prefix)
        if remaining_characters < 0:
            raise _SubstitutionCharacterLimit
        literal = _reference_value_literal(
            reference,
            occurrence.target,
            max_characters=remaining_characters,
        )
        if literal is None:
            return None
        current_length += len(prefix) + len(literal)
        pieces.extend((prefix, literal))
        source_index = occurrence.end

    suffix = cell.formula_a1[source_index:]
    if current_length + len(suffix) > max_characters:
        raise _SubstitutionCharacterLimit
    pieces.append(suffix)
    return "".join(pieces)


def _reference_value_literal(
    reference: ExcelFormulaReference,
    target: FormulaReferenceTarget,
    *,
    max_characters: int,
) -> str | None:
    expected_count = target.row_count * target.column_count
    if len(reference.cells) != expected_count:
        return None
    if max_characters <= 0:
        raise _SubstitutionCharacterLimit

    if expected_count == 1:
        return _scalar_value_literal(
            reference.cells[0], max_characters=max_characters
        )

    pieces = ["{"]
    current_length = 1
    for index, reference_cell in enumerate(reference.cells):
        separator = ""
        if index:
            separator = (
                ";"
                if index % target.column_count == 0
                else ","
            )
        remaining_characters = (
            max_characters - current_length - len(separator) - 1
        )
        if remaining_characters < 0:
            raise _SubstitutionCharacterLimit
        literal = _scalar_value_literal(
            reference_cell,
            max_characters=remaining_characters,
        )
        if literal is None:
            return None
        pieces.extend((separator, literal))
        current_length += len(separator) + len(literal)
    if current_length + 1 > max_characters:
        raise _SubstitutionCharacterLimit
    pieces.append("}")
    return "".join(pieces)


def _scalar_value_literal(
    cell: ExcelReferenceCell,
    *,
    max_characters: int,
) -> str | None:
    if cell.value_kind == "blank" and cell.value is None:
        # Deliberately non-executable marker: a true blank is not equivalent
        # to an empty-string formula result for ISBLANK and similar functions.
        if max_characters < len("BLANK()"):
            raise _SubstitutionCharacterLimit
        return "BLANK()"
    if cell.value is None:
        return None
    if cell.value_kind == "number":
        normalized = normalize_excel_number_text(cell.value)
        try:
            number = Decimal(normalized)
        except InvalidOperation:
            return None
        if not number.is_finite():
            return None
        if len(normalized) > max_characters:
            raise _SubstitutionCharacterLimit
        return normalized
    if cell.value_kind == "boolean":
        lowered = cell.value.casefold()
        if lowered in {"true", "1", "-1"}:
            literal = "TRUE"
            if len(literal) > max_characters:
                raise _SubstitutionCharacterLimit
            return literal
        if lowered in {"false", "0"}:
            literal = "FALSE"
            if len(literal) > max_characters:
                raise _SubstitutionCharacterLimit
            return literal
        return None
    if cell.value_kind == "error":
        if not cell.value.startswith("#"):
            return None
        if len(cell.value) > max_characters:
            raise _SubstitutionCharacterLimit
        return cell.value
    if cell.value_kind == "text":
        # The Excel literal is quotes around the text with every quote doubled.
        # Check the exact derived length before allocating a potentially
        # multi-megabyte replacement string.
        literal_length = 2 + len(cell.value) + cell.value.count('"')
        if literal_length > max_characters:
            raise _SubstitutionCharacterLimit
        return '"' + cell.value.replace('"', '""') + '"'
    return None


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
