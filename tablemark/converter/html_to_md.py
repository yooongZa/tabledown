"""Convert HTML <table> (e.g. from Excel) to Markdown table."""
from bs4 import BeautifulSoup, NavigableString


_BLOCK_TAGS = {
    "p", "div", "section", "article", "header", "footer", "main", "aside",
    "h1", "h2", "h3", "h4", "h5", "h6",
    "ul", "ol", "li", "blockquote", "pre", "table", "tr", "hr",
}


def html_table_to_markdown(html: str, fill_merged_headers: bool = False) -> str:
    """Parse HTML and return a markdown table string.

    Merged cells become blanks (Markdown cannot draw a span). When
    ``fill_merged_headers`` is set, the *header frame* is forward-filled so the
    blanks left by a merge carry their value — a group-header band (a colspan
    label) spreads right, and the left key columns (a rowspan/grouping value)
    spread down. The data/value region is never filled, so a genuinely empty
    value stays empty (the user-controlled "빈칸을 자동 채우기" option). See
    _fill_header_frame.

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

    if fill_merged_headers:
        rows = _fill_header_frame(table, rows)

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


def html_table_to_model(html: str) -> tuple[list[list[str]], list[list[str]]]:
    """Parse an HTML <table> into a ``(header_levels, data_rows)`` model, merge-aware.

    This feeds the XML path, so it handles merged cells the way a record-style
    export wants and *keeps the header hierarchy* instead of flattening it:

    - Vertically merged cells (``rowspan``) are *forward-filled* — the value is
      repeated down its span instead of leaving blanks, so every row is a
      complete record.
    - A full-width merged *title* row at the top (a single cell spanning every
      column) is dropped, so the real header row below becomes the header
      (otherwise the title would become the column names).
    - Multiple header rows (``<thead>``, all-``<th>`` rows, or colspan group
      headers) are returned as separate aligned levels — the XML builder nests
      them as ``<group>`` ("매출" over "1분기"), preserving sub-headers a single
      flat header would collapse.

    When there is no explicit header markup (the common all-``<td>`` Excel
    paste), the first row is treated as the single header level.

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

    header_indices = _detect_header_rows(kept, row_meta)
    data_indices = [index for index in kept if index not in header_indices]

    header_levels = [grid[index] for index in header_indices]
    data_rows = [grid[index] for index in data_indices]
    return _trim_model_columns(header_levels, data_rows)


def html_table_to_rows(html: str) -> list[list[str]]:
    """Flat rows (first row = combined header) derived from ``html_table_to_model``.

    Multi-level headers are joined into composite names ("매출" over "1분기" ->
    "매출 1분기") — the single-header shape the Markdown path and forward-fill
    operate on. The XML path uses ``html_table_to_model`` directly to keep the
    levels nestable.
    """
    header_levels, data_rows = html_table_to_model(html)
    return [_combine_header_levels(header_levels)] + data_rows


def forward_fill_key_columns(rows: list[list[str]]) -> list[list[str]]:
    """Forward-fill blanks in the left *key* columns, vertically then horizontally.

    Many real tables express grouping without merging cells: a grouping column
    (직급, 부서, …) is filled only on the first row of each group, the rest left
    blank meaning "same as above"; a multi-level group may also leave a left
    cell blank where the column above is blank too, meaning "same as the left".
    ``_table_to_filled_grid`` already repeats genuine ``rowspan``/``colspan``
    merges, but a plain blank ``<td>`` stays empty — this fills those from the
    cell directly above, then (for cells still blank) from the cell to the left,
    like a merged cell inheriting its rowspan/colspan origin.

    Scope guard (the convention every data tool follows — fill grouping columns,
    not value columns): only the leading key columns are touched. Walk columns
    left→right and treat everything before the first column with no blanks (the
    data region) as key columns; data/value blanks are left untouched, since an
    empty value there is usually "genuinely none", not a repeat. ``rows[0]`` is
    the header and is never filled. Mutates and returns ``rows``.
    """
    body = rows[1:]
    if not body:
        return rows
    ncols = max((len(row) for row in rows), default=0)
    # Key-column range: left columns up to the first fully-populated (data) one.
    key_cols = 0
    for col in range(ncols):
        cells = [row[col] if col < len(row) else "" for row in body]
        if all(cell.strip() for cell in cells):
            break
        key_cols = col + 1
    # 1) Vertical fill (value above) — the common case, a group down a column.
    for col in range(key_cols):
        last = ""
        for row in body:
            if col >= len(row):
                continue
            if row[col].strip():
                last = row[col]
            elif last:
                row[col] = last
    # 2) Horizontal fill (value to the left) for any key cell still blank — a
    #    left group cell whose column above was also empty inherits leftward.
    for row in body:
        last = ""
        for col in range(key_cols):
            if col >= len(row):
                continue
            if row[col].strip():
                last = row[col]
            elif last:
                row[col] = last
    return rows


def _fill_header_frame(table, rows: list[list[str]]) -> list[list[str]]:
    """Forward-fill the *header frame* of a Markdown grid, leaving values blank.

    Markdown cannot draw a merge, but it can carry the merged value: this fills
    the blanks a merge left behind, restricted to the header frame so no value is
    invented in the data region (the same guard forward_fill_key_columns uses):

    - Horizontal — a group-header *band* (a row carrying a ``colspan`` > 1 label,
      e.g. "1분기" over 1/2/3월) spreads its label right across the band. A
      single full-width cell (a *title* row, not a group header) is left alone.
    - Vertical — the left *key* columns (a ``rowspan`` or blank grouping such as
      "부장" over two rows) spread their value down. The key range is the leading
      columns that still hold a blank in the non-header rows, up to the first
      fully-populated column (the data region) — value blanks past it are left as
      "genuinely empty".

    ``rows`` is the padded/trimmed grid (row count matches the raw table, so the
    span ``meta`` indices align); mutated in place and returned.
    """
    if not rows:
        return rows
    _, meta = _table_to_filled_grid(table)
    nrows = len(rows)
    ncols = max(len(row) for row in rows)
    header_rows = set(_detect_header_rows(list(range(nrows)), meta))

    # 1) Horizontal fill of group-header bands.
    for index in header_rows:
        if index >= len(meta):
            continue
        origins = meta[index]
        is_title = len(origins) == 1 and origins[0][0] >= ncols
        if is_title or not any(colspan > 1 for colspan, _ in origins):
            continue
        last = ""
        for col in range(len(rows[index])):
            if rows[index][col].strip():
                last = rows[index][col]
            elif last:
                rows[index][col] = last

    # 2) Vertical fill of the left key columns.
    body = [index for index in range(nrows) if index not in header_rows]
    key_cols = 0
    for col in range(ncols):
        cells = [rows[index][col] for index in body if col < len(rows[index])]
        if cells and all(cell.strip() for cell in cells):
            break
        key_cols = col + 1
    for col in range(key_cols):
        last = ""
        for index in range(nrows):
            if col >= len(rows[index]):
                continue
            if rows[index][col].strip():
                last = rows[index][col]
            elif last:
                rows[index][col] = last
    return rows


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


def _trim_model_columns(
    header_levels: list[list[str]], data_rows: list[list[str]]
) -> tuple[list[list[str]], list[list[str]]]:
    """Drop right-edge columns empty across every header level and data row."""
    all_rows = header_levels + data_rows
    if not all_rows:
        return header_levels, data_rows
    max_cols = max(len(row) for row in all_rows)
    keep_cols = max_cols
    while keep_cols > 1:
        col_index = keep_cols - 1
        if any(_has_content(row[col_index]) for row in all_rows if col_index < len(row)):
            break
        keep_cols -= 1
    return (
        [row[:keep_cols] for row in header_levels],
        [row[:keep_cols] for row in data_rows],
    )


def _combine_header_levels(header_levels: list[list[str]]) -> list[str]:
    """Join multi-level header labels per column ("매출" / "1분기" -> "매출 1분기")."""
    if not header_levels:
        return []
    ncols = max(len(level) for level in header_levels)
    combined = []
    for col in range(ncols):
        labels: list[str] = []
        for level in header_levels:
            value = (level[col] if col < len(level) else "").strip()
            if value and (not labels or labels[-1] != value):
                labels.append(value)
        combined.append(" ".join(labels))
    return combined


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


def convert_document_tables(html: str, fill_merged_headers: bool = False) -> str:
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
            placeholders[key] = html_table_to_markdown(str(table), fill_merged_headers)
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
