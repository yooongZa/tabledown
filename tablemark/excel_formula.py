"""Read a selected Excel table, including values and formulas, on macOS.

Excel table export is explicit and separate from the clipboard watcher. Reading
Excel therefore requests Automation permission only when the user invokes the
menu command; normal clipboard conversion remains permission-free.
"""
from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Protocol

from AppKit import NSRunningApplication
from Foundation import NSAppleScript

from .converter.formula_export import (
    MAX_FORMULA_CHARACTERS,
    MAX_SNAPSHOT_READS,
    MAX_VALUE_CHARACTERS,
    ExcelFormulaSelection,
    ExcelSelectionCell,
)


EXCEL_BUNDLE_ID = "com.microsoft.Excel"
MAX_SELECTION_CELLS = 10_000
MAX_FORMULA_RECTANGLES = 64
# Bounds one Excel Formula2/R1C1 materialization while keeping a dense 10,000
# cell export responsive (five chunks for a 100x100 range in current Excel).
MAX_FORMULA_CHUNK_CELLS = 2_048
MASK_RESULT_STATUS = "ok_mask_v2"
AREA_RESULT_STATUS = "ok_areas_v2"

EXCEL_NOT_RUNNING = "excel_not_running"
NO_SELECTION = "no_selection"
MULTIPLE_AREAS = "multiple_areas"
TOO_MANY_CELLS = "too_many_cells"
TOO_FRAGMENTED = "too_fragmented"
TOO_MUCH_TEXT = "too_much_text"
NO_FORMULAS = "no_formulas"
SELECTION_CHANGED = "selection_changed"
AUTOMATION_DENIED = "automation_denied"
EXECUTION_FAILED = "execution_failed"
INVALID_RESPONSE = "invalid_response"

KNOWN_ERROR_CODES = frozenset(
    {
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
    }
)


class ExcelFormulaError(RuntimeError):
    """A content-free failure safe to log and map to localized UI text."""

    def __init__(self, code: str):
        self.code = code if code in KNOWN_ERROR_CODES else EXECUTION_FAILED
        super().__init__(self.code)


class ExcelFormulaExecutor(Protocol):
    """Small dependency boundary used by unit tests and the native executor."""

    def is_excel_running(self) -> bool: ...

    def run(self, source: str) -> object: ...


_FORMULA_MASK_HELPERS = f'''
on bitMaskText(rawMask)
    if (class of rawMask) is not list then set rawMask to {{rawMask}}
    set maskText to ""
    repeat with flagReference in rawMask
        if (contents of flagReference) is true then
            set maskText to maskText & "1"
        else
            set maskText to maskText & "0"
        end if
    end repeat
    return maskText
end bitMaskText

on evaluatedFormulaMask(maskRange, expectedCellCount)
    using terms from application "Microsoft Excel"
        tell application id "{EXCEL_BUNDLE_ID}"
            set maskAddress to get address maskRange row absolute true column absolute true reference style A1
            set maskExpression to "TEXTJOIN(" & quote & quote & ",FALSE,IF(ISFORMULA(" & maskAddress & ")," & quote & "1" & quote & "," & quote & "0" & quote & "))"
            set maskText to evaluate name maskExpression as text
        end tell
    end using terms from
    if (length of maskText) is not expectedCellCount then error number -2700
    return maskText
end evaluatedFormulaMask

on compactFormulaMask(targetRange, targetRowCount, targetColumnCount)
    -- Excel's Evaluate API reduces arrays larger than 255 cells to one value.
    -- Evaluate bounded rectangles and concatenate their row-major bit strings.
    set maximumChunkCells to 255
    set combinedMaskText to ""
    set rowOffsetValue to 0
    using terms from application "Microsoft Excel"
        tell application id "{EXCEL_BUNDLE_ID}"
            -- A dense formula range is common and Excel can answer this once
            -- without returning any formula or constant values.
            try
                if (get has formula of targetRange) is true then
                    repeat (targetRowCount * targetColumnCount) times
                        set combinedMaskText to combinedMaskText & "1"
                    end repeat
                    return combinedMaskText
                end if
            end try
            repeat while rowOffsetValue < targetRowCount
                if targetColumnCount <= maximumChunkCells then
                    set chunkRowCount to maximumChunkCells div targetColumnCount
                    set remainingRowCount to targetRowCount - rowOffsetValue
                    if chunkRowCount > remainingRowCount then set chunkRowCount to remainingRowCount
                    set chunkStart to get offset targetRange row offset rowOffsetValue column offset 0
                    set maskChunk to get resize chunkStart row size chunkRowCount column size targetColumnCount
                    set combinedMaskText to combinedMaskText & my evaluatedFormulaMask(maskChunk, chunkRowCount * targetColumnCount)
                    set rowOffsetValue to rowOffsetValue + chunkRowCount
                else
                    set columnOffsetValue to 0
                    repeat while columnOffsetValue < targetColumnCount
                        set chunkColumnCount to targetColumnCount - columnOffsetValue
                        if chunkColumnCount > maximumChunkCells then set chunkColumnCount to maximumChunkCells
                        set chunkStart to get offset targetRange row offset rowOffsetValue column offset columnOffsetValue
                        set maskChunk to get resize chunkStart row size 1 column size chunkColumnCount
                        set combinedMaskText to combinedMaskText & my evaluatedFormulaMask(maskChunk, chunkColumnCount)
                        set columnOffsetValue to columnOffsetValue + chunkColumnCount
                    end repeat
                    set rowOffsetValue to rowOffsetValue + 1
                end if
            end repeat
        end tell
    end using terms from
    if (length of combinedMaskText) is not (targetRowCount * targetColumnCount) then error number -2700
    return combinedMaskText
end compactFormulaMask
'''


# Stage 1 reads metadata plus one content-free IsFormula flag per selected
# cell. Formula text is deliberately absent: Python first turns the mask into
# bounded ranges containing formulas only.
EXCEL_MASK_SCRIPT = f'''
{_FORMULA_MASK_HELPERS}

using terms from application "Microsoft Excel"
    tell application id "{EXCEL_BUNDLE_ID}"
        try
            set selectedRange to selection
            set selectedAreas to get areas of selectedRange
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{NO_SELECTION}"}}
        end try

        if (count of selectedAreas) is not 1 then
            return {{"error", "{MULTIPLE_AREAS}"}}
        end if

        try
            set selectedRowCount to count of rows of selectedRange
            set selectedColumnCount to count of columns of selectedRange
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{NO_SELECTION}"}}
        end try
        if selectedRowCount > {MAX_SELECTION_CELLS} or selectedColumnCount > {MAX_SELECTION_CELLS} then
            return {{"error", "{TOO_MANY_CELLS}"}}
        end if
        if (selectedRowCount * selectedColumnCount) > {MAX_SELECTION_CELLS} then
            return {{"error", "{TOO_MANY_CELLS}"}}
        end if

        try
            set workbookName to name of active workbook as text
            set sheetName to name of active sheet as text
            set selectionAddress to get address selectedRange row absolute true column absolute true reference style A1
            try
                -- ISFORMULA is evaluated inside Excel and returns only a
                -- compact bit mask: no values or formulas cross the boundary.
                set formulaMaskText to my compactFormulaMask(selectedRange, selectedRowCount, selectedColumnCount)
            on error
                -- Older Excel releases without TEXTJOIN keep a content-free,
                -- slower fallback.
                set formulaMask to get has formula of every cell of selectedRange
                set formulaMaskText to my bitMaskText(formulaMask)
            end try
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{EXECUTION_FAILED}"}}
        end try
        if (length of formulaMaskText) is not (selectedRowCount * selectedColumnCount) then
            return {{"error", "{EXECUTION_FAILED}"}}
        end if
        if formulaMaskText does not contain "1" then
            return {{"error", "{NO_FORMULAS}"}}
        end if
        return {{"{MASK_RESULT_STATUS}", workbookName, sheetName, selectionAddress, selectedRowCount as text, selectedColumnCount as text, formulaMaskText}}
    end tell
end using terms from
'''

# Backward-compatible name used by packaging checks and external tests.
EXCEL_FORMULA_SCRIPT = EXCEL_MASK_SCRIPT


def _descriptor_to_python(descriptor):
    """Convert a nested Apple Event list into strings/lists only."""
    item_count = int(descriptor.numberOfItems())
    if item_count:
        return [
            _descriptor_to_python(descriptor.descriptorAtIndex_(index))
            for index in range(1, item_count + 1)
        ]
    value = descriptor.stringValue()
    if value is None:
        raise ExcelFormulaError(INVALID_RESPONSE)
    return str(value)


def _error_number(error_info) -> int | None:
    """Read only the numeric AppleScript error; messages may contain user data."""
    if not error_info:
        return None
    try:
        return int(error_info.get("NSAppleScriptErrorNumber"))
    except (AttributeError, TypeError, ValueError):
        return None


class NSAppleScriptExecutor:
    """Native executor backed by ``NSAppleScript`` and running-app lookup."""

    def is_excel_running(self) -> bool:
        return bool(
            NSRunningApplication.runningApplicationsWithBundleIdentifier_(
                EXCEL_BUNDLE_ID
            )
        )

    def run(self, source: str) -> object:
        script = NSAppleScript.alloc().initWithSource_(source)
        descriptor, error_info = script.executeAndReturnError_(None)
        if error_info is not None:
            # Never inspect or propagate a localized message: it can contain
            # workbook-derived data.
            if _error_number(error_info) == -1743:
                raise ExcelFormulaError(AUTOMATION_DENIED)
            raise ExcelFormulaError(EXECUTION_FAILED)
        if descriptor is None:
            raise ExcelFormulaError(INVALID_RESPONSE)
        return _descriptor_to_python(descriptor)


_A1_AREA_PATTERN = re.compile(
    r"^\$([A-Z]{1,3})\$([1-9][0-9]*)(?::\$([A-Z]{1,3})\$([1-9][0-9]*))?$"
)
_MAX_EXCEL_ROW = 1_048_576
_MAX_EXCEL_COLUMN = 16_384


@dataclass(frozen=True)
class _FormulaChunk:
    address: str
    row_count: int
    column_count: int
    start_row: int
    start_column: int
    end_row: int
    end_column: int


@dataclass(frozen=True)
class _FormulaReadPlan:
    workbook: str
    sheet: str
    address: str
    row_count: int
    column_count: int
    start_row: int
    start_column: int
    mask: tuple[bool, ...]
    value_chunks: tuple[_FormulaChunk, ...]
    chunks: tuple[_FormulaChunk, ...]


def _positive_int(value: object) -> int:
    if not isinstance(value, str) or not value.isascii() or not value.isdigit():
        raise ExcelFormulaError(INVALID_RESPONSE)
    parsed = int(value)
    if parsed < 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    return parsed


def _column_number(name: str) -> int:
    number = 0
    for character in name:
        number = (number * 26) + (ord(character) - ord("A") + 1)
    if not 1 <= number <= _MAX_EXCEL_COLUMN:
        raise ExcelFormulaError(INVALID_RESPONSE)
    return number


def _column_name(number: int) -> str:
    if not 1 <= number <= _MAX_EXCEL_COLUMN:
        raise ExcelFormulaError(INVALID_RESPONSE)
    characters = []
    while number:
        number, remainder = divmod(number - 1, 26)
        characters.append(chr(ord("A") + remainder))
    return "".join(reversed(characters))


def _rectangle_bounds(address: object) -> tuple[int, int, int, int]:
    if not isinstance(address, str):
        raise ExcelFormulaError(INVALID_RESPONSE)
    match = _A1_AREA_PATTERN.fullmatch(address)
    if match is None:
        raise ExcelFormulaError(INVALID_RESPONSE)

    start_column = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_column = _column_number(match.group(3) or match.group(1))
    end_row = int(match.group(4) or match.group(2))
    if start_row > _MAX_EXCEL_ROW or end_row > _MAX_EXCEL_ROW:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if end_row < start_row or end_column < start_column:
        raise ExcelFormulaError(INVALID_RESPONSE)
    return start_row, start_column, end_row, end_column


def _absolute_address(
    start_row: int, start_column: int, end_row: int, end_column: int
) -> str:
    first = f"${_column_name(start_column)}${start_row}"
    last = f"${_column_name(end_column)}${end_row}"
    return first if first == last else f"{first}:{last}"


def _flatten_formula_mask(raw_mask: object, expected_count: int) -> tuple[bool, ...]:
    if (
        isinstance(raw_mask, str)
        and len(raw_mask) == expected_count
        and all(character in "01" for character in raw_mask)
    ):
        return tuple(character == "1" for character in raw_mask)
    if isinstance(raw_mask, str):
        values = [raw_mask] if expected_count == 1 else None
    elif (
        isinstance(raw_mask, (list, tuple))
        and len(raw_mask) == expected_count
        and all(not isinstance(value, (list, tuple)) for value in raw_mask)
    ):
        values = list(raw_mask)
    else:
        values = None
    if values is None or any(value not in ("true", "false") for value in values):
        raise ExcelFormulaError(INVALID_RESPONSE)
    return tuple(value == "true" for value in values)


def _row_spans(mask: tuple[bool, ...], row: int, columns: int):
    """Yield inclusive zero-based true spans for one row of a flat mask."""
    start = None
    offset = row * columns
    for column in range(columns):
        is_formula = mask[offset + column]
        if is_formula and start is None:
            start = column
        if start is not None and (not is_formula or column == columns - 1):
            end = column if is_formula else column - 1
            yield start, end
            start = None


def _formula_rectangles(
    mask: tuple[bool, ...],
    rows: int,
    columns: int,
    selection_start_row: int,
    selection_start_column: int,
) -> list[tuple[int, int, int, int]]:
    """Merge identical horizontal formula spans across consecutive rows."""
    active: dict[tuple[int, int], tuple[int, int]] = {}
    completed: list[tuple[int, int, int, int]] = []

    for row_offset in range(rows):
        absolute_row = selection_start_row + row_offset
        spans = list(_row_spans(mask, row_offset, columns))
        current_keys = set(spans)
        for (start_offset, end_offset), (start_row, end_row) in active.items():
            if (start_offset, end_offset) not in current_keys:
                completed.append(
                    (
                        start_row,
                        selection_start_column + start_offset,
                        end_row,
                        selection_start_column + end_offset,
                    )
                )

        next_active: dict[tuple[int, int], tuple[int, int]] = {}
        for span in spans:
            previous = active.get(span)
            start_row = previous[0] if previous is not None else absolute_row
            next_active[span] = (start_row, absolute_row)
        active = next_active

    for (start_offset, end_offset), (start_row, end_row) in active.items():
        completed.append(
            (
                start_row,
                selection_start_column + start_offset,
                end_row,
                selection_start_column + end_offset,
            )
        )
    completed.sort()
    return completed


def _split_rectangle(
    rectangle: tuple[int, int, int, int]
) -> list[_FormulaChunk]:
    start_row, start_column, end_row, end_column = rectangle
    width = end_column - start_column + 1
    tile_width = min(width, MAX_FORMULA_CHUNK_CELLS)
    tile_height = max(1, MAX_FORMULA_CHUNK_CELLS // tile_width)
    chunks = []
    for chunk_start_row in range(start_row, end_row + 1, tile_height):
        chunk_end_row = min(end_row, chunk_start_row + tile_height - 1)
        for chunk_start_column in range(
            start_column, end_column + 1, tile_width
        ):
            chunk_end_column = min(
                end_column, chunk_start_column + tile_width - 1
            )
            row_count = chunk_end_row - chunk_start_row + 1
            column_count = chunk_end_column - chunk_start_column + 1
            chunks.append(
                _FormulaChunk(
                    address=_absolute_address(
                        chunk_start_row,
                        chunk_start_column,
                        chunk_end_row,
                        chunk_end_column,
                    ),
                    row_count=row_count,
                    column_count=column_count,
                    start_row=chunk_start_row,
                    start_column=chunk_start_column,
                    end_row=chunk_end_row,
                    end_column=chunk_end_column,
                )
            )
    return chunks


def _raise_error_payload(payload: object) -> None:
    if not isinstance(payload, (list, tuple)) or not payload:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if payload[0] != "error":
        return
    if len(payload) != 2 or not isinstance(payload[1], str):
        raise ExcelFormulaError(INVALID_RESPONSE)
    raise ExcelFormulaError(payload[1])


def _parse_mask_result(payload: object) -> _FormulaReadPlan:
    _raise_error_payload(payload)
    if not isinstance(payload, (list, tuple)) or len(payload) != 7:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if payload[0] != MASK_RESULT_STATUS:
        raise ExcelFormulaError(INVALID_RESPONSE)

    workbook, sheet, address, raw_rows, raw_columns, raw_mask = payload[1:]
    if not all(
        isinstance(value, str) and value for value in (workbook, sheet, address)
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)
    start_row, start_column, end_row, end_column = _rectangle_bounds(address)
    rows = _positive_int(raw_rows)
    columns = _positive_int(raw_columns)
    if end_row != start_row + rows - 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if end_column != start_column + columns - 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    cell_count = rows * columns
    if cell_count > MAX_SELECTION_CELLS:
        raise ExcelFormulaError(INVALID_RESPONSE)

    mask = _flatten_formula_mask(raw_mask, cell_count)
    if not any(mask):
        raise ExcelFormulaError(NO_FORMULAS)
    rectangles = _formula_rectangles(mask, rows, columns, start_row, start_column)
    if not rectangles:
        raise ExcelFormulaError(NO_FORMULAS)
    if len(rectangles) > MAX_FORMULA_RECTANGLES:
        raise ExcelFormulaError(TOO_FRAGMENTED)
    chunks = tuple(
        chunk for rectangle in rectangles for chunk in _split_rectangle(rectangle)
    )
    value_chunks = tuple(
        _split_rectangle((start_row, start_column, end_row, end_column))
    )
    return _FormulaReadPlan(
        workbook=workbook,
        sheet=sheet,
        address=address,
        row_count=rows,
        column_count=columns,
        start_row=start_row,
        start_column=start_column,
        mask=mask,
        value_chunks=value_chunks,
        chunks=chunks,
    )


def _codepoint_expression(value: str) -> str:
    """Return an injection-safe AppleScript expression for arbitrary text."""
    return "string id {" + ", ".join(str(ord(character)) for character in value) + "}"


def _chunk_literal(chunk: _FormulaChunk) -> str:
    # Address is generated from validated integers, never from workbook text.
    return f'{{"{chunk.address}", "{chunk.row_count}", "{chunk.column_count}"}}'


def _build_formula_read_script(plan: _FormulaReadPlan) -> str:
    expected_mask_text = "".join("1" if flag else "0" for flag in plan.mask)
    expected_formula_count = sum(plan.mask)
    # Excel's Evaluate command rejects expressions near its legacy 255-character
    # boundary. Keep each group conservative and sum the scalar results in
    # AppleScript; addresses are generated from validated integer coordinates.
    count_expression_groups: list[str] = []
    current_count_expression = ""
    for chunk in plan.chunks:
        term = f"SUMPRODUCT(--ISFORMULA({chunk.address}))"
        candidate = (
            f"{current_count_expression}+{term}"
            if current_count_expression
            else term
        )
        if current_count_expression and len(candidate) > 200:
            count_expression_groups.append(current_count_expression)
            current_count_expression = term
        else:
            current_count_expression = candidate
    count_expression_groups.append(current_count_expression)
    planned_formula_count_literals = ", ".join(
        f'"{expression}"' for expression in count_expression_groups
    )
    # AppleScript does not accept a physical newline after a list comma unless
    # it carries an explicit continuation character. Keep this generated list
    # on one line; at the 10,000-cell limit it remains comfortably bounded.
    chunk_literals = ", ".join(_chunk_literal(chunk) for chunk in plan.chunks)
    value_chunk_literals = ", ".join(
        _chunk_literal(chunk) for chunk in plan.value_chunks
    )
    return f'''
on flattenedText(rawValues)
    set savedTextItemDelimiters to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to ""
        set flattenedValues to rawValues as text
        set AppleScript's text item delimiters to savedTextItemDelimiters
        return flattenedValues
    on error errorMessage number errorNumber
        set AppleScript's text item delimiters to savedTextItemDelimiters
        error number errorNumber
    end try
end flattenedText

on appendFlattenedItems(rawItem, flattenedItems)
    if (class of rawItem) is list then
        repeat with childItemReference in rawItem
            my appendFlattenedItems(contents of childItemReference, flattenedItems)
        end repeat
    else
        set end of flattenedItems to rawItem
    end if
end appendFlattenedItems

on excelErrorText(targetRange, cellIndex)
    using terms from application "Microsoft Excel"
        tell application id "{EXCEL_BUNDLE_ID}"
            try
                set targetCell to get cell cellIndex of targetRange
                set targetAddress to get address targetCell row absolute true column absolute true reference style A1
                set errorCode to (evaluate name ("ERROR.TYPE(" & targetAddress & ")")) as integer
            on error
                return "#ERROR"
            end try
        end tell
    end using terms from
    if errorCode is 1 then return "#NULL!"
    if errorCode is 2 then return "#DIV/0!"
    if errorCode is 3 then return "#VALUE!"
    if errorCode is 4 then return "#REF!"
    if errorCode is 5 then return "#NAME?"
    if errorCode is 6 then return "#NUM!"
    if errorCode is 7 then return "#N/A"
    if errorCode is 8 then return "#GETTING_DATA"
    return "#ERROR"
end excelErrorText

on taggedRangeValues(targetRange, expectedCellCount)
    set rawValues to {{}}
    set rawFormulaFlags to {{}}
    using terms from application "Microsoft Excel"
        tell application id "{EXCEL_BUNDLE_ID}"
            set rawValues to get value of every cell of targetRange
            set rawFormulaFlags to get has formula of every cell of targetRange
        end tell
    end using terms from
    set flatRawValues to {{}}
    set flatFormulaFlags to {{}}
    my appendFlattenedItems(rawValues, flatRawValues)
    my appendFlattenedItems(rawFormulaFlags, flatFormulaFlags)
    if (count of flatRawValues) is not expectedCellCount then error number -2700
    if (count of flatFormulaFlags) is not expectedCellCount then error number -2700

    set taggedValues to {{}}
    repeat with valueIndex from 1 to expectedCellCount
        set rawValue to contents of item valueIndex of flatRawValues
        set isFormulaCell to contents of item valueIndex of flatFormulaFlags
        if isFormulaCell is not true and isFormulaCell is not false then error number -2700
        -- Current Excel returns an empty string, not missing value, for a
        -- genuinely blank cell. Preserve formula results of "" as values.
        if isFormulaCell is false and (rawValue is missing value or ((class of rawValue) is text and rawValue is "")) then
            set end of taggedValues to {{"blank", ""}}
        else
            if rawValue is missing value then
                -- Excel exposes formula errors as missing value through its
                -- AppleScript bridge. ERROR.TYPE recovers a stable error name.
                set valueText to my excelErrorText(targetRange, valueIndex)
            else
                try
                    set valueText to rawValue as text
                on error
                    set valueText to my excelErrorText(targetRange, valueIndex)
                end try
            end if
            set end of taggedValues to {{"value", valueText}}
        end if
    end repeat
    return taggedValues
end taggedRangeValues

{_FORMULA_MASK_HELPERS}

on formulaTopologyMatches(targetRange, targetRowCount, targetColumnCount, expectedFormulaCount, plannedFormulaCountExpressions, expectedMaskText)
    using terms from application "Microsoft Excel"
        tell application id "{EXCEL_BUNDLE_ID}"
            try
                set targetAddress to get address targetRange row absolute true column absolute true reference style A1
                set selectionCountExpression to "SUMPRODUCT(--ISFORMULA(" & targetAddress & "))"
                set selectionFormulaCount to evaluate name selectionCountExpression
                set plannedFormulaCount to 0
                repeat with countExpressionReference in plannedFormulaCountExpressions
                    set countExpression to contents of countExpressionReference
                    set plannedFormulaCount to plannedFormulaCount + (evaluate name countExpression)
                end repeat
                return (selectionFormulaCount is expectedFormulaCount) and (plannedFormulaCount is expectedFormulaCount)
            on error
                -- Compatibility fallback for Excel versions that cannot
                -- evaluate the scalar count expressions.
                try
                    set currentMaskText to my compactFormulaMask(targetRange, targetRowCount, targetColumnCount)
                on error
                    set currentMask to get has formula of every cell of targetRange
                    set currentMaskText to my bitMaskText(currentMask)
                end try
                return currentMaskText is expectedMaskText
            end try
        end tell
    end using terms from
end formulaTopologyMatches

set expectedWorkbookName to {_codepoint_expression(plan.workbook)}
set expectedSheetName to {_codepoint_expression(plan.sheet)}
set expectedSelectionAddress to "{plan.address}"
set expectedMaskText to "{expected_mask_text}"
set expectedFormulaCount to {expected_formula_count}
set plannedFormulaCountExpressions to {{{planned_formula_count_literals}}}
set formulaChunkSpecs to {{{chunk_literals}}}
set valueChunkSpecs to {{{value_chunk_literals}}}

using terms from application "Microsoft Excel"
    tell application id "{EXCEL_BUNDLE_ID}"
        try
            -- Fix object references once. Every formula range below belongs to
            -- this verified sheet, even if UI focus changes during the event.
            set selectedWorkbook to active workbook
            set selectedSheet to active sheet
            set selectedRange to selection
            set selectedAreas to get areas of selectedRange
            if (count of selectedAreas) is not 1 then return {{"error", "{SELECTION_CHANGED}"}}
            if (name of selectedWorkbook as text) is not expectedWorkbookName then return {{"error", "{SELECTION_CHANGED}"}}
            if (name of selectedSheet as text) is not expectedSheetName then return {{"error", "{SELECTION_CHANGED}"}}
            set currentSelectionAddress to get address selectedRange row absolute true column absolute true reference style A1
            if currentSelectionAddress is not expectedSelectionAddress then return {{"error", "{SELECTION_CHANGED}"}}
            if (count of rows of selectedRange) is not {plan.row_count} then return {{"error", "{SELECTION_CHANGED}"}}
            if (count of columns of selectedRange) is not {plan.column_count} then return {{"error", "{SELECTION_CHANGED}"}}
            if (my formulaTopologyMatches(selectedRange, {plan.row_count}, {plan.column_count}, expectedFormulaCount, plannedFormulaCountExpressions, expectedMaskText)) is false then return {{"error", "{SELECTION_CHANGED}"}}
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{SELECTION_CHANGED}"}}
        end try

        set valueAreaPayloads to {{}}
        set valueCharacterCount to 0
        repeat with valueChunkSpecReference in valueChunkSpecs
            set valueChunkSpec to contents of valueChunkSpecReference
            set valueAreaAddress to (item 1 of valueChunkSpec) as text
            set expectedValueRowCount to (item 2 of valueChunkSpec) as integer
            set expectedValueColumnCount to (item 3 of valueChunkSpec) as integer
            try
                set valueArea to range valueAreaAddress of selectedSheet
                if (count of rows of valueArea) is not expectedValueRowCount then return {{"error", "{SELECTION_CHANGED}"}}
                if (count of columns of valueArea) is not expectedValueColumnCount then return {{"error", "{SELECTION_CHANGED}"}}
                set taggedValueValues to my taggedRangeValues(valueArea, expectedValueRowCount * expectedValueColumnCount)
            on error errorMessage number errorNumber
                if errorNumber is -1743 then error number errorNumber
                return {{"error", "{EXECUTION_FAILED}"}}
            end try
            repeat with taggedValueReference in taggedValueValues
                set taggedValue to contents of taggedValueReference
                set valueCharacterCount to valueCharacterCount + (length of ((item 2 of taggedValue) as text))
            end repeat
            if valueCharacterCount > {MAX_VALUE_CHARACTERS} then
                return {{"error", "{TOO_MUCH_TEXT}"}}
            end if
            set end of valueAreaPayloads to {{valueAreaAddress, expectedValueRowCount as text, expectedValueColumnCount as text, taggedValueValues}}
        end repeat

        set formulaAreaPayloads to {{}}
        set formulaCharacterCount to 0
        repeat with formulaChunkSpecReference in formulaChunkSpecs
            set formulaChunkSpec to contents of formulaChunkSpecReference
            -- Coerce list items before passing them back to Excel. Otherwise
            -- Excel treats the AppleScript item reference as a relative range
            -- specifier and can shift B1:C2 to unrelated cells.
            set formulaAreaAddress to (item 1 of formulaChunkSpec) as text
            set expectedRowCount to (item 2 of formulaChunkSpec) as integer
            set expectedColumnCount to (item 3 of formulaChunkSpec) as integer
            try
                set formulaArea to range formulaAreaAddress of selectedSheet
                if (count of rows of formulaArea) is not expectedRowCount then return {{"error", "{SELECTION_CHANGED}"}}
                if (count of columns of formulaArea) is not expectedColumnCount then return {{"error", "{SELECTION_CHANGED}"}}
            on error errorMessage number errorNumber
                if errorNumber is -1743 then error number errorNumber
                return {{"error", "{SELECTION_CHANGED}"}}
            end try

            -- Raw Formula2 codes compile even against older Excel dictionaries;
            -- legacy Formula remains the runtime fallback.
            try
                set formulaA1Values to get «class 2122» of formulaArea
            on error errorMessage number errorNumber
                if errorNumber is -1743 then error number errorNumber
                try
                    set formulaA1Values to get formula of formulaArea
                on error errorMessage number errorNumber
                    if errorNumber is -1743 then error number errorNumber
                    return {{"error", "{EXECUTION_FAILED}"}}
                end try
            end try
            set formulaR1C1State to "present"
            try
                set formulaR1C1Values to get «class F2rc» of formulaArea
            on error errorMessage number errorNumber
                if errorNumber is -1743 then error number errorNumber
                try
                    set formulaR1C1Values to get formula r1c1 of formulaArea
                on error errorMessage number errorNumber
                    if errorNumber is -1743 then error number errorNumber
                    set formulaR1C1State to "missing"
                    set formulaR1C1Values to ""
                end try
            end try
            try
                set formulaCharacterCount to formulaCharacterCount + (length of (my flattenedText(formulaA1Values)))
                if formulaR1C1State is "present" then
                    set formulaCharacterCount to formulaCharacterCount + (length of (my flattenedText(formulaR1C1Values)))
                end if
            on error
                return {{"error", "{EXECUTION_FAILED}"}}
            end try
            if formulaCharacterCount > {MAX_FORMULA_CHARACTERS} then
                return {{"error", "{TOO_MUCH_TEXT}"}}
            end if
            set end of formulaAreaPayloads to {{formulaAreaAddress, expectedRowCount as text, expectedColumnCount as text, formulaA1Values, formulaR1C1State, formulaR1C1Values}}
        end repeat

        -- Recheck after all reads to reject a topology change that happened
        -- during this Apple Event instead of exporting a partial mixed state.
        if (my formulaTopologyMatches(selectedRange, {plan.row_count}, {plan.column_count}, expectedFormulaCount, plannedFormulaCountExpressions, expectedMaskText)) is false then return {{"error", "{SELECTION_CHANGED}"}}

        return {{"{AREA_RESULT_STATUS}", expectedWorkbookName, expectedSheetName, expectedSelectionAddress, "{plan.row_count}", "{plan.column_count}", valueAreaPayloads, formulaAreaPayloads}}
    end tell
end using terms from
'''


def _flatten_formula_matrix(
    raw_values: object, rows: int, columns: int, *, allow_empty: bool
) -> list[str]:
    expected_count = rows * columns
    if isinstance(raw_values, str):
        values = [raw_values] if expected_count == 1 else None
    elif isinstance(raw_values, (list, tuple)):
        if (
            len(raw_values) == rows
            and all(isinstance(row, (list, tuple)) for row in raw_values)
            and all(len(row) == columns for row in raw_values)
        ):
            values = [value for row in raw_values for value in row]
        elif len(raw_values) == expected_count and all(
            not isinstance(value, (list, tuple)) for value in raw_values
        ):
            values = list(raw_values)
        else:
            values = None
    else:
        values = None
    if values is None or not all(isinstance(value, str) for value in values):
        raise ExcelFormulaError(INVALID_RESPONSE)
    if not allow_empty and any(not value for value in values):
        raise ExcelFormulaError(INVALID_RESPONSE)
    return values


def _flatten_tagged_values(
    raw_values: object, rows: int, columns: int
) -> list[str | None]:
    """Decode the stage-2 blank/value tags without conflating ``None`` and ``""``."""
    expected_count = rows * columns
    if not isinstance(raw_values, (list, tuple)) or len(raw_values) != expected_count:
        raise ExcelFormulaError(INVALID_RESPONSE)
    values: list[str | None] = []
    for raw_value in raw_values:
        if not isinstance(raw_value, (list, tuple)) or len(raw_value) != 2:
            raise ExcelFormulaError(INVALID_RESPONSE)
        state, text = raw_value
        if not isinstance(text, str):
            raise ExcelFormulaError(INVALID_RESPONSE)
        if state == "blank" and text == "":
            values.append(None)
        elif state == "value":
            values.append(text)
        else:
            raise ExcelFormulaError(INVALID_RESPONSE)
    return values


def parse_excel_formula_result(
    payload: object, expected_plan: _FormulaReadPlan | None = None
) -> ExcelFormulaSelection:
    """Validate a stage-2 full-range payload and build the shared model."""
    _raise_error_payload(payload)
    if not isinstance(payload, (list, tuple)) or len(payload) != 8:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if payload[0] != AREA_RESULT_STATUS:
        raise ExcelFormulaError(INVALID_RESPONSE)

    (
        workbook,
        sheet,
        address,
        raw_selection_rows,
        raw_selection_columns,
        raw_value_areas,
        raw_formula_areas,
    ) = payload[1:]
    if not all(
        isinstance(value, str) and value for value in (workbook, sheet, address)
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)
    start_row, start_column, end_row, end_column = _rectangle_bounds(address)
    selection_rows = _positive_int(raw_selection_rows)
    selection_columns = _positive_int(raw_selection_columns)
    if end_row != start_row + selection_rows - 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if end_column != start_column + selection_columns - 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if selection_rows * selection_columns > MAX_SELECTION_CELLS:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if expected_plan is not None and (
        workbook != expected_plan.workbook
        or sheet != expected_plan.sheet
        or address != expected_plan.address
        or selection_rows != expected_plan.row_count
        or selection_columns != expected_plan.column_count
    ):
        raise ExcelFormulaError(SELECTION_CHANGED)
    if not isinstance(raw_value_areas, (list, tuple)) or not raw_value_areas:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if not isinstance(raw_formula_areas, (list, tuple)) or not raw_formula_areas:
        raise ExcelFormulaError(NO_FORMULAS)
    if (
        len(raw_value_areas) > MAX_SELECTION_CELLS
        or len(raw_formula_areas) > MAX_SELECTION_CELLS
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)

    expected_value_chunks = None
    expected_formula_chunks = None
    if expected_plan is not None:
        expected_value_chunks = {
            (chunk.address, chunk.row_count, chunk.column_count)
            for chunk in expected_plan.value_chunks
        }
        expected_formula_chunks = {
            (chunk.address, chunk.row_count, chunk.column_count)
            for chunk in expected_plan.chunks
        }

    value_coordinates: dict[tuple[int, int], str | None] = {}
    seen_value_chunks: set[tuple[str, int, int]] = set()
    for raw_area in raw_value_areas:
        if not isinstance(raw_area, (list, tuple)) or len(raw_area) != 4:
            raise ExcelFormulaError(INVALID_RESPONSE)
        area_address, raw_rows, raw_columns, raw_values = raw_area
        rows = _positive_int(raw_rows)
        columns = _positive_int(raw_columns)
        area_start_row, area_start_column, area_end_row, area_end_column = (
            _rectangle_bounds(area_address)
        )
        if area_end_row != area_start_row + rows - 1:
            raise ExcelFormulaError(INVALID_RESPONSE)
        if area_end_column != area_start_column + columns - 1:
            raise ExcelFormulaError(INVALID_RESPONSE)
        if not (
            start_row <= area_start_row <= area_end_row <= end_row
            and start_column <= area_start_column <= area_end_column <= end_column
        ):
            raise ExcelFormulaError(INVALID_RESPONSE)
        chunk_key = (area_address, rows, columns)
        if chunk_key in seen_value_chunks:
            raise ExcelFormulaError(INVALID_RESPONSE)
        seen_value_chunks.add(chunk_key)
        if expected_value_chunks is not None and chunk_key not in expected_value_chunks:
            raise ExcelFormulaError(INVALID_RESPONSE)

        values = _flatten_tagged_values(raw_values, rows, columns)
        for index, value in enumerate(values):
            row_offset, column_offset = divmod(index, columns)
            coordinate = (
                area_start_row + row_offset,
                area_start_column + column_offset,
            )
            if coordinate in value_coordinates:
                raise ExcelFormulaError(INVALID_RESPONSE)
            value_coordinates[coordinate] = value

    all_coordinates = {
        (
            start_row + (index // selection_columns),
            start_column + (index % selection_columns),
        )
        for index in range(selection_rows * selection_columns)
    }
    if value_coordinates.keys() != all_coordinates:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if (
        expected_value_chunks is not None
        and seen_value_chunks != expected_value_chunks
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)
    value_character_count = sum(
        len(value) for value in value_coordinates.values() if value is not None
    )
    if value_character_count > MAX_VALUE_CHARACTERS:
        raise ExcelFormulaError(TOO_MUCH_TEXT)

    formula_character_count = 0
    seen_coordinates: set[tuple[int, int]] = set()
    seen_chunks: set[tuple[str, int, int]] = set()
    coordinate_formulas: dict[tuple[int, int], tuple[str, str | None]] = {}
    for raw_area in raw_formula_areas:
        if not isinstance(raw_area, (list, tuple)) or len(raw_area) != 6:
            raise ExcelFormulaError(INVALID_RESPONSE)
        area_address, raw_rows, raw_columns, raw_a1, r1c1_state, raw_r1c1 = raw_area
        rows = _positive_int(raw_rows)
        columns = _positive_int(raw_columns)
        area_start_row, area_start_column, area_end_row, area_end_column = (
            _rectangle_bounds(area_address)
        )
        if area_end_row != area_start_row + rows - 1:
            raise ExcelFormulaError(INVALID_RESPONSE)
        if area_end_column != area_start_column + columns - 1:
            raise ExcelFormulaError(INVALID_RESPONSE)
        if not (
            start_row <= area_start_row <= area_end_row <= end_row
            and start_column <= area_start_column <= area_end_column <= end_column
        ):
            raise ExcelFormulaError(INVALID_RESPONSE)
        chunk_key = (area_address, rows, columns)
        if chunk_key in seen_chunks:
            raise ExcelFormulaError(INVALID_RESPONSE)
        seen_chunks.add(chunk_key)
        if (
            expected_formula_chunks is not None
            and chunk_key not in expected_formula_chunks
        ):
            raise ExcelFormulaError(INVALID_RESPONSE)

        a1_values = _flatten_formula_matrix(raw_a1, rows, columns, allow_empty=False)
        if r1c1_state == "present":
            r1c1_values = _flatten_formula_matrix(
                raw_r1c1, rows, columns, allow_empty=True
            )
        elif r1c1_state == "missing" and raw_r1c1 == "":
            r1c1_values = [""] * (rows * columns)
        else:
            raise ExcelFormulaError(INVALID_RESPONSE)

        for index, (formula_a1, formula_r1c1) in enumerate(
            zip(a1_values, r1c1_values)
        ):
            row_offset, column_offset = divmod(index, columns)
            cell_row = area_start_row + row_offset
            cell_column = area_start_column + column_offset
            coordinate = (cell_row, cell_column)
            if coordinate in seen_coordinates:
                raise ExcelFormulaError(INVALID_RESPONSE)
            seen_coordinates.add(coordinate)
            formula_character_count += len(formula_a1) + len(formula_r1c1)
            if formula_character_count > MAX_FORMULA_CHARACTERS:
                raise ExcelFormulaError(TOO_MUCH_TEXT)
            coordinate_formulas[coordinate] = (formula_a1, formula_r1c1 or None)

    if (
        expected_formula_chunks is not None
        and seen_chunks != expected_formula_chunks
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)
    if expected_plan is not None:
        expected_coordinates = {
            (
                expected_plan.start_row + (index // expected_plan.column_count),
                expected_plan.start_column + (index % expected_plan.column_count),
            )
            for index, flag in enumerate(expected_plan.mask)
            if flag
        }
        if seen_coordinates != expected_coordinates:
            raise ExcelFormulaError(INVALID_RESPONSE)
    if not coordinate_formulas:
        raise ExcelFormulaError(NO_FORMULAS)

    cells = []
    for row_offset in range(selection_rows):
        for column_offset in range(selection_columns):
            cell_row = start_row + row_offset
            cell_column = start_column + column_offset
            coordinate = (cell_row, cell_column)
            formula_a1, formula_r1c1 = coordinate_formulas.get(
                coordinate, (None, None)
            )
            cells.append(
                ExcelSelectionCell(
                    address=f"${_column_name(cell_column)}${cell_row}",
                    value=value_coordinates[coordinate],
                    formula_a1=formula_a1,
                    formula_r1c1=formula_r1c1,
                )
            )
    return ExcelFormulaSelection(
        workbook=workbook,
        sheet=sheet,
        address=address,
        row_count=selection_rows,
        column_count=selection_columns,
        cells=tuple(cells),
    )


def read_selected_excel_formulas(
    executor: ExcelFormulaExecutor | None = None,
) -> ExcelFormulaSelection:
    """Read every cell from Excel's single rectangular current selection."""
    active_executor = executor or NSAppleScriptExecutor()
    if not active_executor.is_excel_running():
        raise ExcelFormulaError(EXCEL_NOT_RUNNING)
    try:
        mask_payload = active_executor.run(EXCEL_MASK_SCRIPT)
        plan = _parse_mask_result(mask_payload)
        formula_payload = active_executor.run(_build_formula_read_script(plan))
        return parse_excel_formula_result(formula_payload, expected_plan=plan)
    except ExcelFormulaError:
        raise
    except Exception:
        # Do not carry an exception message: it may contain formula content.
        raise ExcelFormulaError(EXECUTION_FAILED) from None


def read_stable_selected_excel_formulas(
    executor: ExcelFormulaExecutor | None = None,
    *,
    max_reads: int = MAX_SNAPSHOT_READS,
) -> ExcelFormulaSelection:
    """Return two matching snapshots, with at most one bounded retry.

    Excel does not expose a transaction around values plus Formula2/R1C1.
    Comparing complete immutable selections catches same-address formula edits
    and recalculation-only value changes that topology checks cannot see. A
    transient ``selection_changed`` consumes one read but remains retryable;
    permission, availability, and execution failures return immediately.
    """
    if max_reads < 2:
        raise ValueError("max_reads must be at least 2")

    previous: ExcelFormulaSelection | None = None
    for attempt in range(max_reads):
        try:
            current = read_selected_excel_formulas(executor)
        except ExcelFormulaError as exc:
            if exc.code == SELECTION_CHANGED and attempt + 1 < max_reads:
                # A failed read breaks consecutiveness. Do not accept the next
                # snapshot merely because it matches one from before the race.
                previous = None
                continue
            raise

        if previous is not None and current == previous:
            return current
        previous = current

    # Content changed on every completed read. Keep formula/value payload out of
    # the exception and let the caller leave the clipboard untouched.
    raise ExcelFormulaError(SELECTION_CHANGED)
