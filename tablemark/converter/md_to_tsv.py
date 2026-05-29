"""Convert Markdown table to spreadsheet-friendly formats."""
from html import escape
import re


def is_markdown_table(text: str, *, strict: bool = True) -> bool:
    """Heuristic check for markdown table format.

    Requires (a) the first line to contain pipes and (b) the second line to be a
    separator of `-`/`:` segments. When ``strict`` (the default), also requires
    (c) the separator to have the same cell count as the header — this rejects
    coincidental `| ... |` text whose second line happens to start with `-`
    (shell output, ASCII art) but does not describe the same columns.

    Callers with independent proof the clipboard holds a real table (e.g. an
    accompanying HTML ``<table>``) pass ``strict=False`` to keep the looser
    check, so a genuine table with a slightly mismatched separator is not
    misclassified as plain text.
    """
    lines = [line.strip() for line in text.strip().split("\n") if line.strip()]
    if len(lines) < 2:
        return False
    if "|" not in lines[0]:
        return False
    if not _is_separator_line(lines[1]):
        return False
    if not strict:
        return True
    return _cell_count(lines[0]) == _cell_count(lines[1])


def _cell_count(line: str) -> int:
    """Count cells in a markdown table row, ignoring leading/trailing pipes."""
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    return len(_split_cells(stripped))


def markdown_table_to_tsv(md: str) -> str:
    """Convert markdown table text to TSV string."""
    rows = markdown_table_to_rows(md)
    return "\n".join("\t".join(row) for row in rows)


def markdown_table_to_html(md: str) -> str:
    """Convert markdown table text to a minimal HTML table."""
    rows = markdown_table_to_rows(md)
    html_rows = []
    for row_index, row in enumerate(rows):
        tag = "th" if row_index == 0 else "td"
        cells = "".join(f"<{tag}>{escape(cell)}</{tag}>" for cell in row)
        html_rows.append(f"<tr>{cells}</tr>")
    return "<table>" + "".join(html_rows) + "</table>"


def markdown_table_to_rows(md: str) -> list[list[str]]:
    """Convert markdown table text to rows of cell values."""
    lines = [line for line in md.strip().split("\n") if line.strip()]
    if not lines:
        raise ValueError("빈 마크다운입니다")

    # Drop separator lines like |---|---|
    data_lines = [line for line in lines if not _is_separator_line(line)]

    rows = []
    for line in data_lines:
        line = line.strip()
        # Strip leading/trailing pipes
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|"):
            line = line[:-1]

        # Split on unescaped pipes
        cells = _split_cells(line)
        cells = [_unescape(c.strip()) for c in cells]
        rows.append(cells)

    return rows


def _is_separator_line(line: str) -> bool:
    """Check if line is a markdown separator like |---|---|."""
    stripped = re.sub(r"[\s|]", "", line)
    if not stripped:
        return False
    return all(c in "-:" for c in stripped)


def _split_cells(line: str) -> list[str]:
    """Split on |, but respect escaped \\| inside cells."""
    # Replace escaped pipe with placeholder, split, then restore
    PLACEHOLDER = "\x00ESCAPED_PIPE\x00"
    line = line.replace("\\|", PLACEHOLDER)
    cells = line.split("|")
    return [c.replace(PLACEHOLDER, "|") for c in cells]


def _unescape(text: str) -> str:
    return text.replace("\\|", "|")
