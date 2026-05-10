"""Convert HTML <table> (e.g. from Excel) to Markdown table."""
from bs4 import BeautifulSoup


def html_table_to_markdown(html: str) -> str:
    """Parse HTML and return a markdown table string.

    Raises:
        ValueError: if no table is found or it's empty.
    """
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        raise ValueError("HTML에 <table>이 없습니다")

    rows = _table_to_grid(table)

    if not rows:
        raise ValueError("표가 비어있습니다")

    # Pad rows to equal column count
    max_cols = max(len(r) for r in rows)
    rows = [r + [" "] * (max_cols - len(r)) for r in rows]
    rows = _trim_trailing_empty_columns(rows)
    max_cols = max(len(r) for r in rows)

    # First row = header. If only one row, treat it as header with empty body.
    header = rows[0]
    body = rows[1:]

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * max_cols) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


def _table_to_grid(table) -> list[list[str]]:
    """Expand rowspan/colspan into a rectangular grid.

    Markdown tables cannot represent merged cells, so merged positions become
    empty cells to preserve row/column alignment.
    """
    rows = []
    occupied = {}

    for row_index, tr in enumerate(table.find_all("tr")):
        row = []
        col_index = 0
        cells = tr.find_all(["td", "th"])
        if not cells and row_index not in occupied:
            continue

        for cell in cells:
            while occupied.pop((row_index, col_index), None) is not None:
                row.append(" ")
                col_index += 1

            rowspan = _span_value(cell.get("rowspan"))
            colspan = _span_value(cell.get("colspan"))
            row.append(_clean_cell(cell.get_text()))

            for offset in range(1, colspan):
                row.append(" ")

            for row_offset in range(1, rowspan):
                for col_offset in range(colspan):
                    occupied[(row_index + row_offset, col_index + col_offset)] = True

            col_index += colspan

        while occupied.pop((row_index, col_index), None) is not None:
            row.append(" ")
            col_index += 1

        rows.append(row)

    if occupied:
        last_row_index = max(row for row, _ in occupied)
        for row_index in range(len(rows), last_row_index + 1):
            row = []
            col_index = 0
            while occupied.pop((row_index, col_index), None) is not None:
                row.append(" ")
                col_index += 1
            rows.append(row)

    return rows


def _trim_trailing_empty_columns(rows: list[list[str]]) -> list[list[str]]:
    """Drop columns at the right edge only when every row is empty there."""
    if not rows:
        return rows

    max_cols = max(len(row) for row in rows)
    keep_cols = max_cols
    while keep_cols > 1:
        col_index = keep_cols - 1
        if any(_has_content(row[col_index]) for row in rows if col_index < len(row)):
            break
        keep_cols -= 1

    return [row[:keep_cols] for row in rows]


def _has_content(cell: str) -> bool:
    return bool(cell.strip())


def _span_value(value) -> int:
    try:
        span = int(value) if value else 1
    except (TypeError, ValueError):
        return 1
    return max(span, 1)


def _clean_cell(text: str) -> str:
    """Escape pipes, collapse whitespace, ensure non-empty."""
    text = text.replace("|", "\\|")
    text = text.replace("\n", " ").replace("\r", " ")
    text = " ".join(text.split())  # collapse whitespace
    return text if text else " "
