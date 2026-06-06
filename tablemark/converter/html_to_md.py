"""Convert HTML <table> (e.g. from Excel) to Markdown table."""
from bs4 import BeautifulSoup, NavigableString


_BLOCK_TAGS = {
    "p", "div", "section", "article", "header", "footer", "main", "aside",
    "h1", "h2", "h3", "h4", "h5", "h6",
    "ul", "ol", "li", "blockquote", "pre", "table", "tr", "hr",
}


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


def html_table_to_rows(html: str) -> list[list[str]]:
    """Parse an HTML <table> into rows (first row = header), merge-aware.

    Unlike ``html_table_to_markdown`` (which targets Markdown's single-header,
    no-merge model), this feeds the XML path, so it handles merged cells the way
    a record-style export wants:

    - Vertically merged cells (``rowspan``) are *forward-filled* — the value is
      repeated down its span instead of leaving blanks, so every row is a
      complete record.
    - A full-width merged *title* row at the top (a single cell spanning every
      column) is dropped, so the real header row below becomes the header
      (otherwise the title would become the column names).
    - Multiple header rows (``<thead>`` or rows made entirely of ``<th>``) are
      combined into composite column names ("매출" over "1분기" -> "매출 1분기"),
      preserving sub-headers that a single Markdown header row would collapse.

    When there is no explicit header markup (the common all-``<td>`` Excel
    paste), the first row is treated as the header, exactly like Markdown.

    Raises:
        ValueError: if no table is found or it's empty.
    """
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        raise ValueError("HTML에 <table>이 없습니다")

    grid, row_meta = _table_to_filled_grid(table)
    if not grid:
        raise ValueError("표가 비어있습니다")

    ncols = max(len(row) for row in grid)
    grid = [row + [""] * (ncols - len(row)) for row in grid]

    # A row whose single origin cell spans the whole width is a title, not data.
    title_rows = {
        index
        for index, origins in enumerate(row_meta)
        if len(origins) == 1 and origins[0][0] >= ncols
    }
    kept = [index for index in range(len(grid)) if index not in title_rows]
    if not kept:
        raise ValueError("표가 비어있습니다")

    header_rows = _detect_header_rows(kept, row_meta)
    data_rows = [index for index in kept if index not in header_rows]

    headers = []
    for col in range(ncols):
        labels = []
        for index in header_rows:
            value = grid[index][col].strip()
            if value and (not labels or labels[-1] != value):
                labels.append(value)
        headers.append(" ".join(labels))

    rows = [headers] + [grid[index] for index in data_rows]
    return _trim_trailing_empty_columns(rows)


def _detect_header_rows(kept: list[int], row_meta: list[list[tuple[int, bool]]]) -> list[int]:
    """Pick which (kept) rows make up the header.

    Priority:
    1. Explicit markup — rows whose every cell is a ``<th>`` / inside ``<thead>``.
    2. Inferred group headers — the common all-``<td>`` Excel export has no ``th``,
       but a group header is given away by a cell spanning multiple columns
       (``colspan``). Count the leading rows that carry such a span, then add the
       one leaf sub-header row beneath them.
    3. Otherwise the single first row, exactly like a Markdown header.

    Never consumes every row — at least one data row is always left.
    """
    explicit = [
        index
        for index in kept
        if row_meta[index] and all(is_header for _, is_header in row_meta[index])
    ]
    if explicit:
        return explicit

    group_rows = 0
    for index in kept:
        if row_meta[index] and any(colspan > 1 for colspan, _ in row_meta[index]):
            group_rows += 1
        else:
            break
    header_count = group_rows + 1 if group_rows >= 1 else 1
    if header_count >= len(kept):
        header_count = 1
    return kept[:header_count]


def _table_to_filled_grid(table) -> tuple[list[list[str]], list[list[tuple[int, bool]]]]:
    """Expand spans into a grid, filling merged positions with the origin value.

    Returns the grid plus, per source row, the ``(colspan, is_header)`` of each
    origin cell (used to spot title rows and header rows). ``rowspan`` repeats
    the value downward and ``colspan`` to the right, so merged cells carry their
    value into every position they cover.
    """
    grid: list[list[str]] = []
    meta: list[list[tuple[int, bool]]] = []
    occupied: dict[tuple[int, int], tuple[str, bool]] = {}

    for row_index, tr in enumerate(table.find_all("tr")):
        cells = tr.find_all(["td", "th"], recursive=False)
        if not cells and not any(key[0] == row_index for key in occupied):
            continue

        is_thead = tr.find_parent("thead") is not None
        row: list[str] = []
        origins: list[tuple[int, bool]] = []
        col = 0

        for cell in cells:
            while (row_index, col) in occupied:
                value, _ = occupied.pop((row_index, col))
                row.append(value)
                col += 1

            rowspan = _span_value(cell.get("rowspan"))
            colspan = _span_value(cell.get("colspan"))
            value = _cell_text_plain(cell)
            is_header = is_thead or cell.name == "th"
            origins.append((colspan, is_header))

            for _ in range(colspan):
                row.append(value)
                for row_offset in range(1, rowspan):
                    occupied[(row_index + row_offset, col)] = (value, is_header)
                col += 1

        while (row_index, col) in occupied:
            value, _ = occupied.pop((row_index, col))
            row.append(value)
            col += 1

        grid.append(row)
        meta.append(origins)

    return grid, meta


def _cell_text_plain(cell) -> str:
    """Cell text for the XML path: collapse whitespace, keep newlines, no pipes.

    Like ``_clean_cell`` but without Markdown's pipe-escaping or <br> joining —
    XML escapes its own special characters and keeps real newlines.
    """
    text = _cell_text(cell).replace("\r\n", "\n").replace("\r", "\n")
    lines = [" ".join(line.split()) for line in text.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
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


def html_has_content_outside_table(html: str) -> bool:
    """True if the HTML carries meaningful content beyond its table(s).

    Excel/Sheets put a bare <table> on the clipboard (plus style/meta noise),
    which is safe to convert to Markdown. A copied *document* (web page, chat
    answer, Word) embeds a table among headings, paragraphs, and lists —
    converting it would discard everything but the table. Detect that so the
    caller can leave the clipboard untouched.
    """
    soup = BeautifulSoup(html, "html.parser")
    for noise in soup(["style", "script", "meta", "title", "head", "link"]):
        noise.decompose()
    for table in soup.find_all("table"):
        table.decompose()
    return bool(soup.get_text(strip=True))


def convert_document_tables(html: str) -> str:
    """Render each <table> in a document as Markdown, keep other text as-is.

    Every <table> becomes a GFM Markdown table; the surrounding content
    (headings, paragraphs, lists) is flattened to plain text — their Markdown
    syntax (`#`, `-`) is NOT added, only tables are converted. Block elements
    are separated by newlines so the result keeps the document's line layout.

    Intended for the text/plain clipboard slot while the original HTML slot is
    left intact: a Markdown editor that reads text sees Markdown tables, while
    rich editors (Word, Excel) still read the original <table> from the HTML
    slot and paste a real table.
    """
    soup = BeautifulSoup(html, "html.parser")
    for noise in soup(["style", "script", "meta", "title", "head", "link"]):
        noise.decompose()

    placeholders = {}
    for index, table in enumerate(soup.find_all("table")):
        key = f"\x00TABLE{index}\x00"
        try:
            placeholders[key] = html_table_to_markdown(str(table))
        except ValueError:
            placeholders[key] = ""
        table.replace_with(key)

    text = _block_text(soup)
    for key, markdown in placeholders.items():
        text = text.replace(key, "\n" + markdown + "\n")
    return _collapse_blank_lines(text)


def _block_text(node) -> str:
    """Extract text, inserting newlines around block-level elements."""
    parts = []
    for child in node.children:
        if isinstance(child, NavigableString):
            parts.append(str(child))
        elif child.name == "br":
            parts.append("\n")
        elif child.name in _BLOCK_TAGS:
            parts.append("\n" + _block_text(child) + "\n")
        else:
            parts.append(_block_text(child))
    return "".join(parts)


def _collapse_blank_lines(text: str) -> str:
    """Strip each line and collapse runs of blank lines to a single one."""
    lines = [line.strip() for line in text.split("\n")]
    out = []
    for line in lines:
        if not line and (not out or not out[-1]):
            continue
        out.append(line)
    while out and not out[-1]:
        out.pop()
    while out and not out[0]:
        out.pop(0)
    return "\n".join(out)
