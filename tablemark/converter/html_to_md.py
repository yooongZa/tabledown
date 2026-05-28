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
            row.append(_clean_cell(_cell_text(cell)))

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


def _cell_text(cell) -> str:
    """Extract cell text, rendering <br> as a literal newline.

    BeautifulSoup's get_text() drops <br> entirely, which would collapse
    Excel/Sheets multi-line cells into a single line. Mutating the cell in
    place is safe because each cell is visited once during grid expansion.
    """
    for br in cell.find_all("br"):
        br.replace_with("\n")
    return cell.get_text()


def _clean_cell(text: str) -> str:
    """Escape pipes, preserve in-cell line breaks as <br>, ensure non-empty.

    Excel Alt+Enter and Sheets Ctrl+Enter put a literal newline inside a cell.
    Markdown table cells cannot contain a raw newline, but GFM (Obsidian,
    GitHub) renders <br> inside a cell as a line break.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    # Collapse runs of spaces/tabs within a line, then join lines with <br>.
    lines = [" ".join(line.split()) for line in text.split("\n")]
    # Drop empty leading/trailing lines so wrap-only whitespace doesn't add
    # phantom breaks, but keep blank lines between content as a single <br>.
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    text = "<br>".join(lines)
    return text if text else " "
