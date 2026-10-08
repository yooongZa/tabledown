"""Read displayed Excel values and merges without touching the clipboard.

The immutable model and XML serializer are shared with macOS. COM access is
read-only, bounded to 10,000 cells, and succeeds only after two equal snapshots.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from tablemark.converter.formula_export import MAX_SNAPSHOT_READS, MAX_VALUE_CHARACTERS
from tablemark.converter.html_to_md import forward_fill_key_columns
from tablemark.converter.table_xml import model_to_xml
from tablemark.excel_formula import ExcelFormulaError, _rectangle_bounds
from tablemark.excel_table import (
    ExcelTableSelection,
    excel_table_selection_to_model_with_sources,
    excel_table_selection_xml_metadata,
)
from . import excel_formula as formulas


class _TableReadError(Exception):
    """Known Windows reader status; never contains workbook contents."""
    def __init__(self, code: str):
        self.code = code
        super().__init__(code)


@dataclass(frozen=True)
class ExcelTableResult:
    code: str
    selection: ExcelTableSelection | None = None

    @property
    def ok(self) -> bool:
        return self.code == formulas.SUCCESS and self.selection is not None


def read_selected_excel_table(
    *, is_cancelled: Callable[[], bool] | None = None,
) -> ExcelTableResult:
    initialized = False
    try:
        pythoncom, client = formulas._load_com_modules()
        pythoncom.CoInitialize()
        initialized = True
        return ExcelTableResult(formulas.SUCCESS, _read_initialized(client, is_cancelled))
    except (ExcelFormulaError, _TableReadError) as exc:
        return ExcelTableResult(exc.code)
    except formulas._FormulaTextTooLarge:
        return ExcelTableResult(formulas.TOO_MUCH_TEXT)
    except formulas._FormulaSnapshotChanged:
        return ExcelTableResult(formulas.SELECTION_CHANGED)
    except Exception:
        # COM messages can contain workbook values; return only a known code.
        return ExcelTableResult(formulas.COM_FAILURE)
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def read_stable_selected_excel_table(
    *, max_reads: int = MAX_SNAPSHOT_READS,
    is_cancelled: Callable[[], bool] | None = None,
) -> ExcelTableResult:
    if max_reads < 2:
        raise ValueError("max_reads must be at least 2")
    previous = None
    for attempt in range(max_reads):
        result = read_selected_excel_table(is_cancelled=is_cancelled)
        if not result.ok:
            if result.code == formulas.SELECTION_CHANGED and attempt + 1 < max_reads:
                previous = None
                continue
            return result
        if result.selection == previous:
            return result
        previous = result.selection
    return ExcelTableResult(formulas.SELECTION_CHANGED)


def _read_initialized(client, is_cancelled) -> ExcelTableSelection:
    def check_cancelled():
        if is_cancelled is not None and is_cancelled():
            raise _TableReadError("clipboard_changed")

    check_cancelled()
    pids = formulas._visible_excel_process_ids()
    if not pids:
        raise _TableReadError(formulas.EXCEL_NOT_RUNNING)
    if len(pids) != 1:
        raise _TableReadError(formulas.MULTIPLE_INSTANCES)
    try:
        excel = client.GetActiveObject("Excel.Application")
    except Exception:
        raise _TableReadError(formulas.EXCEL_NOT_RUNNING) from None
    if formulas._window_process_id(excel.Hwnd) not in pids:
        raise _TableReadError(formulas.MULTIPLE_INSTANCES)
    selection = excel.Selection
    try:
        areas = formulas._count(selection.Areas)
    except Exception:
        raise _TableReadError(formulas.SELECTION_NOT_RANGE) from None
    if areas != 1:
        raise _TableReadError(formulas.MULTIPLE_AREAS)
    identity = formulas._selection_identity(selection)
    workbook, sheet, address, count, _, rows, columns, start_row, start_column = identity
    if count > formulas.MAX_SELECTION_CELLS:
        raise _TableReadError(formulas.TOO_LARGE)
    if rows < 1 or columns < 1 or rows * columns != count:
        raise _TableReadError(formulas.SELECTION_NOT_RANGE)
    raw = formulas._read_bulk_values(
        selection, row_count=rows, column_count=columns,
        max_value_characters=MAX_VALUE_CHARACTERS,
    )
    # Excel returns Null for mixed contents/formats. A scalar Text allows a
    # uniform range to be read once, including the common 10,000-cell boundary.
    uniform_text = selection.Text
    has_merges = selection.MergeCells not in (False, 0)
    values = []
    merges = set()
    covered = set()
    characters = 0
    cells = iter(selection.Cells) if uniform_text is None or has_merges else None
    for index, raw_value in enumerate(raw):
        if index % 32 == 0:
            check_cancelled()
        cell = next(cells) if cells is not None else None
        row, column = start_row + index // columns, start_column + index % columns
        if has_merges and (row, column) not in covered:
            merge = str(cell.MergeArea.Address)
            r1, c1, r2, c2 = _rectangle_bounds(merge)
            if r1 != r2 or c1 != c2:
                if (r1 < start_row or c1 < start_column
                        or r2 >= start_row + rows or c2 >= start_column + columns):
                    raise _TableReadError("partial_merge")
                merges.add(merge)
                covered.update((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
        if raw_value.kind == "blank":
            value = None
        else:
            text = uniform_text if uniform_text is not None else cell.Text
            if not isinstance(text, str):
                raise _TableReadError(formulas.COM_FAILURE)
            if text and set(text) == {"#"} and raw_value.kind != "text":
                raise _TableReadError("display_overflow")
            value = text
        characters += len(value or "")
        if characters > MAX_VALUE_CHARACTERS:
            raise _TableReadError(formulas.TOO_MUCH_TEXT)
        values.append(value)
    check_cancelled()
    if formulas._selection_identity(excel.Selection) != identity:
        raise _TableReadError(formulas.SELECTION_CHANGED)
    return ExcelTableSelection(
        workbook, sheet, address, rows, columns, tuple(values),
        tuple(sorted(merges, key=_rectangle_bounds)),
    )


def selection_to_xml(selection: ExcelTableSelection, *, fill_blanks: bool) -> str:
    headers, rows, sources = excel_table_selection_to_model_with_sources(selection)
    if not rows:
        raise ExcelFormulaError("no_table")
    filled = ()
    if fill_blanks:
        original = [list(row) for row in rows]
        rows = forward_fill_key_columns([list(headers[-1]), *[list(row) for row in rows]])[1:]
        filled = tuple(dict.fromkeys(
            sources[r][c]
            for r, (before, after) in enumerate(zip(original, rows, strict=True))
            for c, (old, new) in enumerate(zip(before, after, strict=True))
            if old != new
        ))
    metadata = excel_table_selection_xml_metadata(
        selection, headers, blank_fill_enabled=fill_blanks, blank_fill_cells=filled,
    )
    return model_to_xml(headers, rows, metadata=metadata)
