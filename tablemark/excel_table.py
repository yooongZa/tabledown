"""Read the current Excel selection as a merge-aware value table on macOS.

The explicit general XML action uses this module instead of requiring the user
to copy first.  It reads values and merge rectangles without touching the
clipboard, then returns two matching immutable snapshots before the caller is
allowed to replace the clipboard with XML.
"""
from __future__ import annotations

from dataclasses import dataclass
from html import escape
from typing import Protocol

from .converter.formula_export import MAX_SNAPSHOT_READS, MAX_VALUE_CHARACTERS
from .converter.html_to_md import html_table_to_model
from .excel_formula import (
    EXCEL_BUNDLE_ID,
    EXCEL_NOT_RUNNING,
    DISPLAY_OVERFLOW,
    EXECUTION_FAILED,
    INVALID_RESPONSE,
    MAX_SELECTION_CELLS,
    MULTIPLE_AREAS,
    NO_SELECTION,
    PARTIAL_MERGE,
    SELECTION_CHANGED,
    TOO_MANY_CELLS,
    TOO_MUCH_TEXT,
    ExcelFormulaError,
    NSAppleScriptExecutor,
    _positive_int,
    _raise_error_payload,
    _rectangle_bounds,
)


TABLE_RESULT_STATUS = "ok_table_v1"
_VALUE_DELIMITER = "\ue000"
_VALUE_ESCAPE = "\ue001"
# General and formula XML use the same public selection bound.  Merge topology
# remains lossless at this size because the script asks Excel for all merge-area
# addresses in two batched Apple Events instead of querying merged cells one by
# one.  Keep the alias explicit so tests and UI copy cannot drift between the
# two direct-selection commands.
MAX_TABLE_SELECTION_CELLS = MAX_SELECTION_CELLS
# A malformed range-level raw-value response can happen when Excel errors are
# mixed into a selection.  Preserve a few literal-hash cells with targeted
# reads, but never regress into an unbounded cell-by-cell Apple Event loop.
MAX_HASH_LITERAL_FALLBACK_CELLS = 32


class ExcelTableExecutor(Protocol):
    def is_excel_running(self) -> bool: ...

    def run(self, source: str) -> object: ...


@dataclass(frozen=True)
class ExcelTableSelection:
    workbook: str
    sheet: str
    address: str
    row_count: int
    column_count: int
    values: tuple[str | None, ...]
    merge_areas: tuple[str, ...]


EXCEL_TABLE_SCRIPT = f'''
on appendFlattenedItems(rawItem, flattenedItems)
    if (class of rawItem) is list then
        repeat with childItemReference in rawItem
            my appendFlattenedItems(contents of childItemReference, flattenedItems)
        end repeat
    else
        set end of flattenedItems to rawItem
    end if
end appendFlattenedItems

on replaceText(sourceText, findText, replacementText)
    set savedDelimiters to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to findText
        set textParts to text items of sourceText
        set AppleScript's text item delimiters to replacementText
        set replacedText to textParts as text
    on error
        set AppleScript's text item delimiters to savedDelimiters
        error number -2700
    end try
    set AppleScript's text item delimiters to savedDelimiters
    return replacedText
end replaceText

on joinValueTexts(valueTexts)
    set encodedValues to {{}}
    repeat with valueTextReference in valueTexts
        set encodedValue to contents of valueTextReference
        if encodedValue contains "{_VALUE_ESCAPE}" then set encodedValue to my replaceText(encodedValue, "{_VALUE_ESCAPE}", "{_VALUE_ESCAPE}e")
        if encodedValue contains "{_VALUE_DELIMITER}" then set encodedValue to my replaceText(encodedValue, "{_VALUE_DELIMITER}", "{_VALUE_ESCAPE}d")
        set end of encodedValues to encodedValue
    end repeat
    set savedDelimiters to AppleScript's text item delimiters
    try
        set AppleScript's text item delimiters to "{_VALUE_DELIMITER}"
        set valueBlob to encodedValues as text
    on error
        set AppleScript's text item delimiters to savedDelimiters
        error number -2700
    end try
    set AppleScript's text item delimiters to savedDelimiters
    return valueBlob
end joinValueTexts

on isAllHashes(candidateText)
    if candidateText is "" then return false
    repeat with characterReference in characters of candidateText
        if (contents of characterReference) is not "#" then return false
    end repeat
    return true
end isAllHashes

on taggedRangeValues(targetRange, targetRowCount, targetColumnCount)
    set expectedCellCount to targetRowCount * targetColumnCount
    set rawStringValues to {{}}
    using terms from application "Microsoft Excel"
        tell application id "{EXCEL_BUNDLE_ID}"
            -- ``string value`` is Excel's exact cell text for text/whitespace
            -- and displayed text for numbers, dates, percentages, currency,
            -- custom formats, booleans, and current error values.
            -- Query the rectangular range itself. Asking for the property of
            -- ``every cell`` makes Excel materialize thousands of object
            -- specifiers and took ~130 seconds for 4,008 cells; the range-level
            -- properties return the same 2-D values in one fast Apple Event.
            set rawStringValues to get string value of targetRange
        end tell
    end using terms from
    set flatStringValues to {{}}
    my appendFlattenedItems(rawStringValues, flatStringValues)
    if (count of flatStringValues) is not expectedCellCount then error number -2700

    set displayedTexts to {{}}
    set hashCandidateIndexes to {{}}
    repeat with valueIndex from 1 to expectedCellCount
        set displayedValue to contents of item valueIndex of flatStringValues
        if displayedValue is missing value then error number -2700
        try
            set displayedText to displayedValue as text
        on error
            error number -2700
        end try
        set end of displayedTexts to displayedText
        if my isAllHashes(displayedText) then set end of hashCandidateIndexes to valueIndex
    end repeat

    if (count of hashCandidateIndexes) > 0 then
        -- ``string value`` returns only #### for a formatted number/date that
        -- cannot be displayed. Distinguish real text consisting of hashes with
        -- one range-level raw-value Apple Event, not one event per candidate.
        -- Excel can return a malformed raw range shape when error cells are
        -- present; only then use a strictly bounded targeted fallback.
        set flatRawValues to {{}}
        set completeRawBatch to false
        try
            using terms from application "Microsoft Excel"
                tell application id "{EXCEL_BUNDLE_ID}"
                    set rawValues to get value of targetRange
                end tell
            end using terms from
            my appendFlattenedItems(rawValues, flatRawValues)
            if (count of flatRawValues) is expectedCellCount then set completeRawBatch to true
        end try

        if completeRawBatch then
            repeat with hashIndexReference in hashCandidateIndexes
                set hashIndex to contents of hashIndexReference
                set rawHashValue to contents of item hashIndex of flatRawValues
                set displayedHashText to contents of item hashIndex of displayedTexts
                if (class of rawHashValue) is not text then return {{false, "", {{}}}}
                if (rawHashValue as text) is not displayedHashText then return {{false, "", {{}}}}
            end repeat
        else
            if (count of hashCandidateIndexes) > {MAX_HASH_LITERAL_FALLBACK_CELLS} then return {{false, "", {{}}}}
            repeat with hashIndexReference in hashCandidateIndexes
                set hashIndex to contents of hashIndexReference
                set displayedHashText to contents of item hashIndex of displayedTexts
                using terms from application "Microsoft Excel"
                    tell application id "{EXCEL_BUNDLE_ID}"
                        try
                            set rawHashValue to get value of cell hashIndex of targetRange
                        on error
                            return {{false, "", {{}}}}
                        end try
                    end tell
                end using terms from
                if (class of rawHashValue) is not text then return {{false, "", {{}}}}
                if (rawHashValue as text) is not displayedHashText then return {{false, "", {{}}}}
            end repeat
        end if
    end if

    -- Keep the large Apple Event response flat: one compact kind mask plus one
    -- list of strings. A nested {{kind, value}} pair per cell multiplies native
    -- descriptor decoding cost on selections with thousands of cells.
    set valueKinds to ""
    set valueTexts to {{}}
    repeat with displayedTextReference in displayedTexts
        set displayedText to contents of displayedTextReference
        if displayedText is "" then
            set valueKinds to valueKinds & "b"
            set end of valueTexts to ""
        else
            set valueKinds to valueKinds & "v"
            set end of valueTexts to displayedText
        end if
    end repeat
    return {{true, valueKinds, valueTexts}}
end taggedRangeValues

using terms from application "Microsoft Excel"
    tell application id "{EXCEL_BUNDLE_ID}"
        try
            set selectedWorkbook to active workbook
            set selectedSheet to active sheet
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
            set selectedCellCount to selectedRowCount * selectedColumnCount
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{NO_SELECTION}"}}
        end try
        if selectedRowCount > {MAX_TABLE_SELECTION_CELLS} or selectedColumnCount > {MAX_TABLE_SELECTION_CELLS} then
            return {{"error", "{TOO_MANY_CELLS}"}}
        end if
        if selectedCellCount > {MAX_TABLE_SELECTION_CELLS} then
            return {{"error", "{TOO_MANY_CELLS}"}}
        end if

        try
            set workbookName to name of selectedWorkbook as text
            set sheetName to name of selectedSheet as text
            set selectionAddress to get address selectedRange row absolute true column absolute true reference style A1
            set compactValues to my taggedRangeValues(selectedRange, selectedRowCount, selectedColumnCount)
            if item 1 of compactValues is false then
                return {{"error", "{DISPLAY_OVERFLOW}"}}
            end if
            set valueKinds to item 2 of compactValues
            set valueTexts to item 3 of compactValues

            set valueCharacterCount to 0
            repeat with valueTextReference in valueTexts
                set valueCharacterCount to valueCharacterCount + (length of ((contents of valueTextReference) as text))
            end repeat
            if valueCharacterCount > {MAX_VALUE_CHARACTERS} then
                return {{"error", "{TOO_MUCH_TEXT}"}}
            end if
            set valueBlob to my joinValueTexts(valueTexts)

            -- A mixed range reports ``merge cells = false`` even when some of
            -- its cells are merged. Excel can instead return every cell's
            -- merge area in one batch; unmerged cells resolve to their own
            -- one-cell address and are filtered out below. This keeps exact
            -- merge topology without one slow Apple Event per merged cell.
            set rawMergeAreas to get merge area of every cell of selectedRange
            set rawMergeAddresses to get address of every item of rawMergeAreas row absolute true column absolute true reference style A1
            if (count of rawMergeAddresses) is not selectedCellCount then error number -2700
            set mergedAddresses to {{}}
            repeat with rawMergeAddressReference in rawMergeAddresses
                set targetMergeAddress to contents of rawMergeAddressReference
                if (class of targetMergeAddress) is not text then error number -2700
                -- Leave duplicate real-merge addresses for Python's set-based
                -- canonicalization, but omit the thousands of singleton
                -- addresses produced by unmerged cells from the Apple Event.
                if targetMergeAddress contains ":" then set end of mergedAddresses to targetMergeAddress
            end repeat
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{EXECUTION_FAILED}"}}
        end try

        try
            set currentWorkbookName to name of active workbook as text
            set currentSheetName to name of active sheet as text
            set currentRange to selection
            set currentAreas to get areas of currentRange
            if (count of currentAreas) is not 1 then return {{"error", "{SELECTION_CHANGED}"}}
            set currentAddress to get address currentRange row absolute true column absolute true reference style A1
            if currentWorkbookName is not workbookName then return {{"error", "{SELECTION_CHANGED}"}}
            if currentSheetName is not sheetName then return {{"error", "{SELECTION_CHANGED}"}}
            if currentAddress is not selectionAddress then return {{"error", "{SELECTION_CHANGED}"}}
            if (count of rows of currentRange) is not selectedRowCount then return {{"error", "{SELECTION_CHANGED}"}}
            if (count of columns of currentRange) is not selectedColumnCount then return {{"error", "{SELECTION_CHANGED}"}}
        on error errorMessage number errorNumber
            if errorNumber is -1743 then error number errorNumber
            return {{"error", "{SELECTION_CHANGED}"}}
        end try

        return {{"{TABLE_RESULT_STATUS}", workbookName, sheetName, selectionAddress, selectedRowCount as text, selectedColumnCount as text, valueKinds, valueBlob, mergedAddresses}}
    end tell
end using terms from
'''


def parse_excel_table_result(payload: object) -> ExcelTableSelection:
    """Validate an Apple Event payload without exposing workbook cell text."""
    _raise_error_payload(payload)
    if not isinstance(payload, (list, tuple)) or len(payload) != 9:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if payload[0] != TABLE_RESULT_STATUS:
        raise ExcelFormulaError(INVALID_RESPONSE)

    (
        workbook,
        sheet,
        address,
        raw_rows,
        raw_columns,
        raw_value_kinds,
        raw_value_blob,
        raw_merges,
    ) = payload[1:]
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
    if rows * columns > MAX_TABLE_SELECTION_CELLS:
        raise ExcelFormulaError(INVALID_RESPONSE)

    expected_cells = rows * columns
    if (
        not isinstance(raw_value_kinds, str)
        or len(raw_value_kinds) != expected_cells
        or any(kind not in "bv" for kind in raw_value_kinds)
        or not isinstance(raw_value_blob, str)
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)
    raw_value_texts = _decode_value_blob(raw_value_blob, expected_cells)
    if any(
        kind == "b" and value != ""
        for kind, value in zip(raw_value_kinds, raw_value_texts, strict=True)
    ):
        raise ExcelFormulaError(INVALID_RESPONSE)
    values = tuple(
        None if kind == "b" else value
        for kind, value in zip(raw_value_kinds, raw_value_texts, strict=True)
    )
    if sum(len(value) for value in values if value is not None) > MAX_VALUE_CHARACTERS:
        raise ExcelFormulaError(TOO_MUCH_TEXT)
    if not isinstance(raw_merges, (list, tuple)):
        raise ExcelFormulaError(INVALID_RESPONSE)

    merge_records: list[tuple[int, int, int, int, str]] = []
    seen_addresses: set[str] = set()
    occupied: set[tuple[int, int]] = set()
    for raw_merge in raw_merges:
        if not isinstance(raw_merge, str):
            raise ExcelFormulaError(INVALID_RESPONSE)
        merge_start_row, merge_start_column, merge_end_row, merge_end_column = (
            _rectangle_bounds(raw_merge)
        )
        # The batch result contains one address per selected cell. Unmerged
        # cells resolve to their own one-cell ranges; every cell in a real merge
        # repeats that merge's address. Filter/canonicalize in Python with a set
        # instead of AppleScript's O(n²) list membership checks.
        if merge_start_row == merge_end_row and merge_start_column == merge_end_column:
            continue
        if raw_merge in seen_addresses:
            continue
        if (
            merge_start_row < start_row
            or merge_start_column < start_column
            or merge_end_row > end_row
            or merge_end_column > end_column
        ):
            raise ExcelFormulaError(PARTIAL_MERGE)
        coordinates = {
            (row, column)
            for row in range(merge_start_row, merge_end_row + 1)
            for column in range(merge_start_column, merge_end_column + 1)
        }
        if occupied.intersection(coordinates):
            raise ExcelFormulaError(INVALID_RESPONSE)
        occupied.update(coordinates)
        seen_addresses.add(raw_merge)
        merge_records.append(
            (
                merge_start_row,
                merge_start_column,
                merge_end_row,
                merge_end_column,
                raw_merge,
            )
        )

    merge_records.sort()
    return ExcelTableSelection(
        workbook=workbook,
        sheet=sheet,
        address=address,
        row_count=rows,
        column_count=columns,
        values=tuple(values),
        merge_areas=tuple(record[-1] for record in merge_records),
    )


def _decode_value_blob(blob: str, expected_count: int) -> tuple[str, ...]:
    """Decode the escaped, delimiter-joined cell strings from AppleScript."""
    values: list[str] = []
    current: list[str] = []
    index = 0
    while index < len(blob):
        character = blob[index]
        if character == _VALUE_DELIMITER:
            values.append("".join(current))
            current = []
        elif character == _VALUE_ESCAPE:
            index += 1
            if index >= len(blob):
                raise ExcelFormulaError(INVALID_RESPONSE)
            escaped = blob[index]
            if escaped == "e":
                current.append(_VALUE_ESCAPE)
            elif escaped == "d":
                current.append(_VALUE_DELIMITER)
            else:
                raise ExcelFormulaError(INVALID_RESPONSE)
        else:
            current.append(character)
        index += 1
    values.append("".join(current))
    if len(values) != expected_count:
        raise ExcelFormulaError(INVALID_RESPONSE)
    return tuple(values)


def read_selected_excel_table(
    executor: ExcelTableExecutor | None = None,
) -> ExcelTableSelection:
    """Read one rectangular selection without requiring any formula cells."""
    active_executor = executor or NSAppleScriptExecutor()
    if not active_executor.is_excel_running():
        raise ExcelFormulaError(EXCEL_NOT_RUNNING)
    try:
        return parse_excel_table_result(active_executor.run(EXCEL_TABLE_SCRIPT))
    except ExcelFormulaError:
        raise
    except Exception:
        raise ExcelFormulaError(EXECUTION_FAILED) from None


def read_stable_selected_excel_table(
    executor: ExcelTableExecutor | None = None,
    *,
    max_reads: int = MAX_SNAPSHOT_READS,
) -> ExcelTableSelection:
    """Return two matching value+merge snapshots, with one bounded retry."""
    if max_reads < 2:
        raise ValueError("max_reads must be at least 2")

    previous: ExcelTableSelection | None = None
    for attempt in range(max_reads):
        try:
            current = read_selected_excel_table(executor)
        except ExcelFormulaError as exc:
            if exc.code == SELECTION_CHANGED and attempt + 1 < max_reads:
                previous = None
                continue
            raise
        if previous is not None and current == previous:
            return current
        previous = current
    raise ExcelFormulaError(SELECTION_CHANGED)


def excel_table_selection_to_html(selection: ExcelTableSelection) -> str:
    """Build one merge-aware HTML table for the existing XML model parser."""
    start_row, start_column, end_row, end_column = _rectangle_bounds(
        selection.address
    )
    if end_row != start_row + selection.row_count - 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if end_column != start_column + selection.column_count - 1:
        raise ExcelFormulaError(INVALID_RESPONSE)
    if len(selection.values) != selection.row_count * selection.column_count:
        raise ExcelFormulaError(INVALID_RESPONSE)

    merge_starts: dict[tuple[int, int], tuple[int, int]] = {}
    covered: set[tuple[int, int]] = set()
    for address in selection.merge_areas:
        merge_start_row, merge_start_column, merge_end_row, merge_end_column = (
            _rectangle_bounds(address)
        )
        if (
            merge_start_row < start_row
            or merge_start_column < start_column
            or merge_end_row > end_row
            or merge_end_column > end_column
        ):
            raise ExcelFormulaError(PARTIAL_MERGE)
        coordinates = {
            (row, column)
            for row in range(merge_start_row, merge_end_row + 1)
            for column in range(merge_start_column, merge_end_column + 1)
        }
        if covered.intersection(coordinates):
            raise ExcelFormulaError(INVALID_RESPONSE)
        covered.update(coordinates)
        merge_starts[(merge_start_row, merge_start_column)] = (
            merge_end_row - merge_start_row + 1,
            merge_end_column - merge_start_column + 1,
        )

    html_rows = []
    for row_offset in range(selection.row_count):
        absolute_row = start_row + row_offset
        cells = []
        for column_offset in range(selection.column_count):
            absolute_column = start_column + column_offset
            coordinate = (absolute_row, absolute_column)
            merge_shape = merge_starts.get(coordinate)
            if coordinate in covered and merge_shape is None:
                continue
            value = selection.values[
                row_offset * selection.column_count + column_offset
            ]
            text = "" if value is None else value
            escaped = escape(text, quote=False).replace("\r\n", "\n")
            escaped = escaped.replace("\r", "\n").replace("\n", "<br>")
            attributes = ""
            if merge_shape is not None:
                rowspan, colspan = merge_shape
                if rowspan > 1:
                    attributes += f' rowspan="{rowspan}"'
                if colspan > 1:
                    attributes += f' colspan="{colspan}"'
            cells.append(f"<td{attributes}>{escaped}</td>")
        html_rows.append("<tr>" + "".join(cells) + "</tr>")
    return "<table>" + "".join(html_rows) + "</table>"


def excel_table_selection_to_model(
    selection: ExcelTableSelection,
) -> tuple[list[list[str]], list[list[str]]]:
    """Infer the merge/header structure without normalizing Excel cell text.

    The existing HTML parser is the canonical source of title removal,
    rowspan/colspan expansion, and multi-level header inference. Feeding real
    values through it would collapse significant whitespace. Instead every
    selected cell is represented by an opaque token during structural parsing,
    then the exact original string is restored into the resulting model.
    """
    tokens = tuple(f"TABLEDOWN_DIRECT_CELL_{index:08X}" for index in range(len(selection.values)))
    token_selection = ExcelTableSelection(
        workbook=selection.workbook,
        sheet=selection.sheet,
        address=selection.address,
        row_count=selection.row_count,
        column_count=selection.column_count,
        values=tokens,
        merge_areas=selection.merge_areas,
    )
    token_headers, token_rows = html_table_to_model(
        excel_table_selection_to_html(token_selection)
    )
    originals = {
        token: "" if value is None else value
        for token, value in zip(tokens, selection.values, strict=True)
    }

    def restore(rows: list[list[str]]) -> list[list[str]]:
        return [[originals.get(value, value) for value in row] for row in rows]

    return restore(token_headers), restore(token_rows)
