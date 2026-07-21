"""Read an Excel table, including values and formulas, through COM.

The COM imports intentionally stay inside :func:`_load_com_modules`.  Besides
keeping non-Windows imports harmless, this makes it explicit that COM is only
entered by the tray action's private worker thread.
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes
from dataclasses import dataclass, replace
from typing import Any, Iterator

from tablemark.converter.formula_export import (
    MAX_FORMULA_CHARACTERS,
    MAX_REFERENCE_CELLS,
    MAX_REFERENCE_RANGES,
    MAX_SNAPSHOT_READS,
    MAX_VALUE_CHARACTERS,
    ExcelFormulaReference,
    ExcelFormulaSelection,
    ExcelReferenceCell,
    ExcelSelectionCell,
    extract_formula_reference_targets,
)


MAX_SELECTION_CELLS = 10_000
MAX_REFERENCE_RANGE_CELLS = 2_048
XL_CELL_TYPE_FORMULAS = -4123

SUCCESS = "success"
EXCEL_NOT_RUNNING = "excel_not_running"
SELECTION_NOT_RANGE = "selection_not_range"
MULTIPLE_AREAS = "multiple_areas"
NO_FORMULAS = "no_formulas"
TOO_LARGE = "too_large"
COM_FAILURE = "com_failure"
MULTIPLE_INSTANCES = "multiple_instances"
TOO_MUCH_TEXT = "too_much_text"
SELECTION_CHANGED = "selection_changed"

_EXCEL_WINDOW_CLASS = "XLMAIN"
_WINDOW_CLASS_BUFFER_SIZE = 256
_EXCEL_ERROR_HRESULT_PREFIX = 0x800A0000
_EXCEL_ERROR_TEXT_BY_CODE = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
    2043: "#GETTING_DATA",
    2045: "#SPILL!",
}


class _FormulaTextTooLarge(Exception):
    """Internal control flow; never carries formula or cell-value text."""


class _FormulaSnapshotChanged(Exception):
    """Internal fail-closed signal; never carries workbook or formula data."""


@dataclass(frozen=True)
class _FormulaMetadata:
    """Formula-only metadata captured before it is joined to bulk values."""

    address: str
    formula_a1: str
    formula_r1c1: str | None


@dataclass(frozen=True)
class ExcelFormulaResult:
    """Result of reading the active Excel formula selection.

    Error results deliberately carry only a stable code.  In particular, no
    exception string is retained because COM errors can include workbook or
    cell content that must not reach the diagnostic log.
    """

    code: str
    selection: ExcelFormulaSelection | None = None

    @property
    def ok(self) -> bool:
        return self.code == SUCCESS and self.selection is not None


def read_selected_excel_formulas(
    *,
    max_cells: int = MAX_SELECTION_CELLS,
    max_formula_characters: int = MAX_FORMULA_CHARACTERS,
    max_value_characters: int = MAX_VALUE_CHARACTERS,
) -> ExcelFormulaResult:
    """Read every cell from one active Excel Range, including formulas.

    ``CoInitialize``/``CoUninitialize`` are paired here rather than in the tray
    app so every caller gets correct thread-local COM setup.  The tray invokes
    this function from a daemon worker and therefore never blocks its Win32
    message pump.
    """

    try:
        pythoncom, win32com_client = _load_com_modules()
    except Exception:  # noqa: BLE001 - dependency/load failures are user-safe codes
        return _error(COM_FAILURE)

    initialized = False
    try:
        pythoncom.CoInitialize()
        initialized = True
    except Exception:  # noqa: BLE001 - do not retain the potentially sensitive error
        return _error(COM_FAILURE)

    try:
        return _read_initialized(
            win32com_client,
            max_cells=max_cells,
            max_formula_characters=max_formula_characters,
            max_value_characters=max_value_characters,
        )
    except Exception:  # noqa: BLE001 - the adapter boundary never leaks COM details
        return _error(COM_FAILURE)
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                # There is no useful recovery at process/thread teardown, and
                # retaining the COM exception would violate the payload rule.
                pass


def read_stable_selected_excel_formulas(
    *,
    max_cells: int = MAX_SELECTION_CELLS,
    max_formula_characters: int = MAX_FORMULA_CHARACTERS,
    max_value_characters: int = MAX_VALUE_CHARACTERS,
    max_reads: int = MAX_SNAPSHOT_READS,
) -> ExcelFormulaResult:
    """Return two matching snapshots, with at most one bounded retry."""
    if max_reads < 2:
        raise ValueError("max_reads must be at least 2")

    previous: ExcelFormulaSelection | None = None
    for attempt in range(max_reads):
        result = read_selected_excel_formulas(
            max_cells=max_cells,
            max_formula_characters=max_formula_characters,
            max_value_characters=max_value_characters,
        )
        if not result.ok:
            if result.code == SELECTION_CHANGED and attempt + 1 < max_reads:
                # A transient race breaks the consecutive comparison window.
                previous = None
                continue
            return result

        current = result.selection
        if previous is not None and current == previous:
            return result
        previous = current

    # Exact formula/value payloads changed on every read. Return only a stable,
    # content-free code so logs and dialogs cannot disclose workbook data.
    return _error(SELECTION_CHANGED)


def _read_initialized(
    win32com_client: Any,
    *,
    max_cells: int,
    max_formula_characters: int,
    max_value_characters: int,
) -> ExcelFormulaResult:
    try:
        visible_excel_pids = _visible_excel_process_ids()
    except Exception:  # noqa: BLE001 - fail closed when instance identity is uncertain
        return _error(COM_FAILURE)
    if not visible_excel_pids:
        # A selected range is only meaningful in a visible Excel window.  Do
        # not let GetActiveObject silently bind to a hidden automation instance.
        return _error(EXCEL_NOT_RUNNING)
    if len(visible_excel_pids) > 1:
        return _error(MULTIPLE_INSTANCES)

    try:
        excel = win32com_client.GetActiveObject("Excel.Application")
    except Exception:  # noqa: BLE001 - no running/accessible Excel instance
        return _error(EXCEL_NOT_RUNNING)

    # GetActiveObject may bind to a different Excel process than the window the
    # user can see.  With exactly one visible Excel PID, require the returned
    # Application.Hwnd to belong to it before reading any workbook content.
    try:
        active_pid = _window_process_id(excel.Hwnd)
    except Exception:  # noqa: BLE001 - an unverifiable identity must fail closed
        return _error(COM_FAILURE)
    if active_pid not in visible_excel_pids:
        return _error(MULTIPLE_INSTANCES)

    try:
        selection = excel.Selection
    except Exception:  # noqa: BLE001 - Excel exists, but querying it failed
        return _error(COM_FAILURE)
    if selection is None:
        return _error(SELECTION_NOT_RANGE)

    # Range is the Excel selection type that exposes Areas.  A chart, shape, or
    # other object selection does not, so a failure at this narrow type probe is
    # reported as "select cells" rather than as a generic COM error.
    try:
        areas = selection.Areas
        area_count = _count(areas)
    except Exception:  # noqa: BLE001 - non-Range selections fail this probe
        return _error(SELECTION_NOT_RANGE)
    if area_count < 1:
        return _error(SELECTION_NOT_RANGE)
    if area_count != 1:
        return _error(MULTIPLE_AREAS)

    try:
        selection_count = _range_count(selection)
    except Exception:  # noqa: BLE001 - valid Range whose properties became unavailable
        return _error(COM_FAILURE)
    if selection_count > max_cells:
        return _error(TOO_LARGE)

    single_cell_formula = False
    if selection_count == 1:
        # Range.SpecialCells has a documented/observed single-cell quirk: it
        # can search the entire UsedRange.  HasFormula answers the only question
        # needed here without ever producing an out-of-selection candidate.
        try:
            has_formula = selection.HasFormula
        except Exception:  # noqa: BLE001 - scope-safe determination unavailable
            return _error(COM_FAILURE)
        if has_formula in (False, 0):
            return _error(NO_FORMULAS)
        if has_formula not in (True, -1):
            return _error(COM_FAILURE)
        single_cell_formula = True

    try:
        worksheet = selection.Worksheet
        workbook = worksheet.Parent
        workbook_name = str(workbook.Name)
        sheet_name = str(worksheet.Name)
        range_address = str(selection.Address)
        row_count = _range_count(selection.Rows)
        column_count = _range_count(selection.Columns)
        start_row = int(selection.Row)
        start_column = int(selection.Column)
    except Exception:  # noqa: BLE001 - metadata is required for a complete export
        return _error(COM_FAILURE)
    if (
        row_count < 1
        or column_count < 1
        or start_row < 1
        or start_column < 1
        or row_count * column_count != selection_count
    ):
        return _error(COM_FAILURE)
    selection_identity = (
        workbook_name,
        sheet_name,
        range_address,
        selection_count,
        area_count,
        row_count,
        column_count,
        start_row,
        start_column,
    )

    if single_cell_formula:
        formula_range = selection
    else:
        # SpecialCells is important for a multi-cell selection: reading formula
        # properties from the full selection would emit constants too.  Excel
        # raises when no matching cells exist.
        try:
            formula_candidates = selection.SpecialCells(XL_CELL_TYPE_FORMULAS)
        except Exception:  # noqa: BLE001 - "No cells were found" is a COM error
            return _error(NO_FORMULAS)
        if formula_candidates is None:
            return _error(NO_FORMULAS)

        # Keep a defensive intersection even for multi-cell ranges.  No formula
        # property is read until Excel confirms the original selection boundary.
        try:
            formula_range = excel.Intersect(formula_candidates, selection)
        except Exception:  # noqa: BLE001 - scope cannot be verified, fail closed
            return _error(COM_FAILURE)
        if formula_range is None:
            return _error(NO_FORMULAS)

    try:
        formula_cells = tuple(
            _read_formula_cells(
                formula_range, max_formula_characters=max_formula_characters
            )
        )
    except _FormulaTextTooLarge:
        return _error(TOO_MUCH_TEXT)
    except _FormulaSnapshotChanged:
        return _error(SELECTION_CHANGED)
    except Exception:  # noqa: BLE001 - do not expose formula-bearing COM exceptions
        return _error(COM_FAILURE)
    if not formula_cells:
        return _error(NO_FORMULAS)

    # Value2 is deliberately read once from the complete rectangular Range.
    # This is both substantially faster than one COM round-trip per cell and
    # ensures formula cells carry their calculated result alongside constants.
    try:
        values = _read_bulk_values(
            selection,
            row_count=row_count,
            column_count=column_count,
            max_value_characters=max_value_characters,
        )
        cells = _join_selection_cells(
            values,
            formula_cells=formula_cells,
            row_count=row_count,
            column_count=column_count,
            start_row=start_row,
            start_column=start_column,
        )
    except _FormulaTextTooLarge:
        return _error(TOO_MUCH_TEXT)
    except _FormulaSnapshotChanged:
        return _error(SELECTION_CHANGED)
    except Exception:  # noqa: BLE001 - never retain value-bearing exceptions
        return _error(COM_FAILURE)

    selection_payload = ExcelFormulaSelection(
        workbook=workbook_name,
        sheet=sheet_name,
        address=range_address,
        row_count=row_count,
        column_count=column_count,
        cells=cells,
    )
    selection_payload = _enrich_formula_references(
        workbook,
        selection_payload,
        max_value_characters=max_value_characters,
    )

    # Re-read the active selection after all formula properties.  Any movement,
    # workbook/sheet switch, range change, or formula-address change invalidates
    # the snapshot.  Only addresses are inspected here; no second content copy
    # is created and no identity detail reaches an error payload or log.
    try:
        current_selection = excel.Selection
        if _selection_identity(current_selection) != selection_identity:
            raise _FormulaSnapshotChanged
        final_topology = _formula_topology(
            excel, current_selection, selection_count=selection_count
        )
    except _FormulaSnapshotChanged:
        return _error(SELECTION_CHANGED)
    except Exception:  # noqa: BLE001 - every inconsistency is content-free failure
        return _error(COM_FAILURE)
    initial_topology = tuple(sorted(cell.address for cell in formula_cells))
    if final_topology != initial_topology:
        return _error(SELECTION_CHANGED)

    return ExcelFormulaResult(SUCCESS, selection_payload)


def _enrich_formula_references(
    workbook: Any,
    selection: ExcelFormulaSelection,
    *,
    max_value_characters: int,
) -> ExcelFormulaSelection:
    """Attach bounded static A1 reference values without changing Excel UI state."""

    references_by_owner: dict[str, list[ExcelFormulaReference]] = {}
    completeness: dict[str, bool] = {}
    reference_range_count = 0
    reference_cell_count = 0
    value_character_count = sum(
        len(cell.value) for cell in selection.cells if cell.value is not None
    )

    for cell in selection.cells:
        if cell.formula_a1 is None:
            continue
        targets, complete = extract_formula_reference_targets(
            cell.formula_a1, selection.sheet
        )
        completeness[cell.address] = complete
        for target in targets:
            target_cell_count = target.row_count * target.column_count
            if (
                reference_range_count >= MAX_REFERENCE_RANGES
                or target_cell_count > MAX_REFERENCE_RANGE_CELLS
                or reference_cell_count + target_cell_count > MAX_REFERENCE_CELLS
            ):
                completeness[cell.address] = False
                continue
            reference_range_count += 1
            reference_cell_count += target_cell_count
            try:
                worksheet = _worksheet_by_name(workbook, target.sheet)
                target_range = _range_by_address(worksheet, target.address)
                if (
                    _range_count(target_range) != target_cell_count
                    or _range_count(target_range.Rows) != target.row_count
                    or _range_count(target_range.Columns) != target.column_count
                    or str(target_range.Address) != target.address
                ):
                    raise _FormulaSnapshotChanged
                values = _read_bulk_values(
                    target_range,
                    row_count=target.row_count,
                    column_count=target.column_count,
                    max_value_characters=max_value_characters,
                )
            except Exception:  # noqa: BLE001 - enrichment never exposes COM detail
                completeness[cell.address] = False
                continue

            candidate_characters = value_character_count + sum(
                len(value) for value in values if value is not None
            )
            if candidate_characters > max_value_characters:
                completeness[cell.address] = False
                continue
            value_character_count = candidate_characters
            start_row = int(target_range.Row)
            start_column = int(target_range.Column)
            reference_cells = tuple(
                ExcelReferenceCell(
                    _absolute_address(
                        start_row + (index // target.column_count),
                        start_column + (index % target.column_count),
                    ),
                    value,
                )
                for index, value in enumerate(values)
            )
            references_by_owner.setdefault(cell.address, []).append(
                ExcelFormulaReference(
                    sheet=target.sheet,
                    address=target.address,
                    cells=reference_cells,
                )
            )

    return replace(
        selection,
        cells=tuple(
            replace(
                cell,
                references=tuple(references_by_owner.get(cell.address, ())),
                references_complete=completeness.get(cell.address, True),
            )
            if cell.formula_a1 is not None
            else cell
            for cell in selection.cells
        ),
    )


def _worksheet_by_name(workbook: Any, name: str) -> Any:
    worksheets = workbook.Worksheets
    item = worksheets.Item
    return item(name) if callable(item) else item[name]


def _range_by_address(worksheet: Any, address: str) -> Any:
    range_member = worksheet.Range
    return range_member(address) if callable(range_member) else range_member[address]


def _read_formula_cells(
    formula_range: Any, *, max_formula_characters: int
) -> Iterator[_FormulaMetadata]:
    """Yield every cell from every area returned by ``SpecialCells``.

    A formula-only result can itself be discontiguous even though the original
    selection is one area.  Iterating each formula area avoids Excel's
    first-area-only behavior for ``Range.Cells`` on a multi-area range.
    """

    formula_characters = 0
    for cell in _range_cells(formula_range):
        address = str(cell.Address)
        _require_formula_cell(cell)
        formula_a1 = _formula_value(cell, "Formula2", "Formula", required=True)
        _require_formula_cell(cell)
        formula_characters += len(formula_a1)
        if formula_characters > max_formula_characters:
            raise _FormulaTextTooLarge

        formula_r1c1 = _formula_value(
            cell, "Formula2R1C1", "FormulaR1C1", required=False
        )
        _require_formula_cell(cell)
        formula_characters += len(formula_r1c1 or "")
        if formula_characters > max_formula_characters:
            # Abort while walking COM cells.  Do not finish materializing a
            # payload that the common serializer will reject anyway.
            raise _FormulaTextTooLarge
        yield _FormulaMetadata(
            address=address,
            formula_a1=formula_a1,
            formula_r1c1=formula_r1c1,
        )


def _read_bulk_values(
    selection: Any,
    *,
    row_count: int,
    column_count: int,
    max_value_characters: int,
) -> tuple[str | None, ...]:
    """Return one bulk ``Range.Value2`` snapshot flattened row-major."""

    raw_values = selection.Value2
    expected_count = row_count * column_count
    if expected_count == 1:
        flattened = (raw_values,)
    else:
        if not isinstance(raw_values, (list, tuple)) or len(raw_values) != row_count:
            raise _FormulaSnapshotChanged
        flattened_rows: list[Any] = []
        for raw_row in raw_values:
            if not isinstance(raw_row, (list, tuple)) or len(raw_row) != column_count:
                raise _FormulaSnapshotChanged
            flattened_rows.extend(raw_row)
        flattened = tuple(flattened_rows)

    if len(flattened) != expected_count:
        raise _FormulaSnapshotChanged

    value_characters = 0
    values: list[str | None] = []
    for value in flattened:
        text = _value_to_text(value)
        value_characters += len(text or "")
        if value_characters > max_value_characters:
            raise _FormulaTextTooLarge
        values.append(text)
    return tuple(values)


def _value_to_text(value: Any) -> str | None:
    """Convert one Value2 result, including Excel CVErr HRESULTs, to text."""

    if value is None:
        return None
    if type(value) is int:
        unsigned = value & 0xFFFFFFFF
        if unsigned & 0xFFFF0000 == _EXCEL_ERROR_HRESULT_PREFIX:
            return _EXCEL_ERROR_TEXT_BY_CODE.get(unsigned & 0xFFFF, "#ERROR")
    return str(value)


def _join_selection_cells(
    values: tuple[str | None, ...],
    *,
    formula_cells: tuple[_FormulaMetadata, ...],
    row_count: int,
    column_count: int,
    start_row: int,
    start_column: int,
) -> tuple[ExcelSelectionCell, ...]:
    """Join bulk values and formula metadata into a complete row-major table."""

    formulas_by_address: dict[str, _FormulaMetadata] = {}
    for formula in formula_cells:
        if formula.address in formulas_by_address:
            raise _FormulaSnapshotChanged
        formulas_by_address[formula.address] = formula

    cells: list[ExcelSelectionCell] = []
    for index, value in enumerate(values):
        row_offset, column_offset = divmod(index, column_count)
        address = _absolute_address(
            start_row + row_offset, start_column + column_offset
        )
        formula = formulas_by_address.pop(address, None)
        cells.append(
            ExcelSelectionCell(
                address=address,
                value=value,
                formula_a1=formula.formula_a1 if formula is not None else None,
                formula_r1c1=(
                    formula.formula_r1c1 if formula is not None else None
                ),
            )
        )
    if formulas_by_address or len(cells) != row_count * column_count:
        raise _FormulaSnapshotChanged
    return tuple(cells)


def _absolute_address(row: int, column: int) -> str:
    """Return an absolute A1 address for a positive row and column."""

    if row < 1 or column < 1:
        raise _FormulaSnapshotChanged
    letters = ""
    remaining = column
    while remaining:
        remaining, remainder = divmod(remaining - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"${letters}${row}"


def _range_cells(formula_range: Any) -> Iterator[Any]:
    formula_areas = formula_range.Areas
    for area_index in range(1, _count(formula_areas) + 1):
        area = _item(formula_areas, area_index)
        area_cells = area.Cells
        for cell_index in range(1, _range_count(area_cells) + 1):
            yield _item(area_cells, cell_index)


def _require_formula_cell(cell: Any) -> None:
    try:
        has_formula = cell.HasFormula
    except Exception:  # noqa: BLE001 - convert COM detail to empty signal
        raise _FormulaSnapshotChanged from None
    if has_formula not in (True, -1):
        raise _FormulaSnapshotChanged


def _selection_identity(
    selection: Any,
) -> tuple[str, str, str, int, int, int, int, int, int]:
    areas = selection.Areas
    worksheet = selection.Worksheet
    return (
        str(worksheet.Parent.Name),
        str(worksheet.Name),
        str(selection.Address),
        _range_count(selection),
        _count(areas),
        _range_count(selection.Rows),
        _range_count(selection.Columns),
        int(selection.Row),
        int(selection.Column),
    )


def _formula_topology(
    excel: Any, selection: Any, *, selection_count: int
) -> tuple[str, ...]:
    if selection_count == 1:
        _require_formula_cell(selection)
        formula_range = selection
    else:
        formula_candidates = selection.SpecialCells(XL_CELL_TYPE_FORMULAS)
        if formula_candidates is None:
            raise _FormulaSnapshotChanged
        formula_range = excel.Intersect(formula_candidates, selection)
        if formula_range is None:
            raise _FormulaSnapshotChanged

    addresses = []
    for cell in _range_cells(formula_range):
        _require_formula_cell(cell)
        address = str(cell.Address)
        _require_formula_cell(cell)
        addresses.append(address)
    return tuple(sorted(addresses))


def _formula_value(
    cell: Any, primary: str, fallback: str, *, required: bool
) -> str | None:
    for property_name in (primary, fallback):
        try:
            value = getattr(cell, property_name)
        except Exception:  # noqa: BLE001 - old Excel lacks Formula2 properties
            continue
        if value is not None:
            text = str(value)
            if text:
                return text
    if required:
        raise RuntimeError("Excel formula property unavailable")
    return None


def _range_count(value: Any) -> int:
    """Return CountLarge, falling back to Count for small/mocked ranges."""

    try:
        return int(value.CountLarge)
    except (AttributeError, TypeError, ValueError):
        return int(value.Count)


def _count(collection: Any) -> int:
    return int(collection.Count)


def _item(collection: Any, index: int) -> Any:
    item = collection.Item
    return item(index) if callable(item) else item[index]


def _visible_excel_process_ids(user32: Any | None = None) -> set[int]:
    """Return unique PIDs for visible top-level ``XLMAIN`` windows.

    ``EnumWindows`` already limits enumeration to top-level windows.  A PID set
    (rather than a window count) intentionally allows one Excel process with
    several workbook windows.
    """

    if user32 is None:
        user32 = _load_user32()
    if user32 is None:
        return set()

    pids: set[int] = set()
    callback_failed = False
    callback_factory = getattr(ctypes, "WINFUNCTYPE", ctypes.CFUNCTYPE)
    enum_proc_type = callback_factory(
        wintypes.BOOL, wintypes.HWND, wintypes.LPARAM
    )

    @enum_proc_type
    def visit_window(hwnd, _lparam):
        nonlocal callback_failed
        try:
            handle = ctypes.c_void_p(int(hwnd))
            if not user32.IsWindowVisible(handle):
                return True

            class_name = ctypes.create_unicode_buffer(_WINDOW_CLASS_BUFFER_SIZE)
            if user32.GetClassNameW(
                handle, class_name, _WINDOW_CLASS_BUFFER_SIZE
            ) <= 0:
                return True
            if class_name.value.upper() != _EXCEL_WINDOW_CLASS:
                return True

            pid = ctypes.c_ulong(0)
            if user32.GetWindowThreadProcessId(handle, ctypes.byref(pid)) == 0:
                callback_failed = True
                return False
            if pid.value:
                pids.add(int(pid.value))
            return True
        except Exception:  # noqa: BLE001 - callback exceptions cannot cross ctypes
            callback_failed = True
            return False

    completed = user32.EnumWindows(visit_window, 0)
    if not completed or callback_failed:
        raise OSError("Unable to enumerate Excel windows")
    return pids


def _window_process_id(hwnd: Any, user32: Any | None = None) -> int:
    if user32 is None:
        user32 = _load_user32()
    if user32 is None:
        raise OSError("Windows user32 is unavailable")

    pid = ctypes.c_ulong(0)
    handle = ctypes.c_void_p(int(hwnd))
    if user32.GetWindowThreadProcessId(handle, ctypes.byref(pid)) == 0 or not pid.value:
        raise OSError("Unable to identify Excel window process")
    return int(pid.value)


def _load_user32() -> Any | None:
    windll = getattr(ctypes, "windll", None)
    return getattr(windll, "user32", None) if windll is not None else None


def _error(code: str) -> ExcelFormulaResult:
    return ExcelFormulaResult(code=code, selection=None)


def _load_com_modules() -> tuple[Any, Any]:
    # Deferred by design: importing this adapter remains safe on macOS/Linux,
    # and COM is loaded only when the explicit Windows tray action is invoked.
    import pythoncom
    import win32com.client

    return pythoncom, win32com.client
