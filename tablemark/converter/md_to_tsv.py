"""Convert Markdown table to spreadsheet-friendly formats."""
from html import escape, unescape
import re
from string import punctuation


_CELL_ESCAPE = re.compile(
    r"\\[" + re.escape(punctuation) + r"]"
    r"|<br\s*/?>"
    r"|&(?:\#[xX][0-9a-fA-F]+|\#[0-9]+|[a-zA-Z][a-zA-Z0-9]+);",
    re.IGNORECASE,
)


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
    return len(_split_cells(_strip_outer_pipes(line)))


def markdown_table_to_tsv(md: str) -> str:
    """Convert markdown table text to TSV string."""
    rows = markdown_table_to_rows(md)
    return "\n".join("\t".join(row) for row in rows)


def markdown_table_to_html(md: str) -> str:
    """Convert Markdown to a table with an explicit clipboard HTML encoding."""
    rows = markdown_table_to_rows(md)
    html_rows = []
    for row_index, row in enumerate(rows):
        tag = "th" if row_index == 0 else "td"
        escaped_cells = (escape(cell).replace("\n", "<br>") for cell in row)
        cells = "".join(f"<{tag}>{cell}</{tag}>" for cell in escaped_cells)
        html_rows.append(f"<tr>{cells}</tr>")
    # Rich-text importers such as Apple Notes can guess a legacy encoding for
    # UTF-8 clipboard bytes unless the generated HTML declares its charset.
    return '<meta charset="utf-8"><table>' + "".join(html_rows) + "</table>"


def markdown_table_to_rows(md: str) -> list[list[str]]:
    """Convert markdown table text to rows of cell values."""
    lines = [line for line in md.strip().split("\n") if line.strip()]
    if not lines:
        raise ValueError("빈 마크다운입니다")

    # Only the second line is the Markdown header separator. A later row made
    # entirely of values such as "-" / "--" is data and must be preserved.
    data_lines = list(lines)
    if len(data_lines) > 1 and _is_separator_line(data_lines[1]):
        del data_lines[1]

    rows = []
    for line in data_lines:
        # Split on unescaped pipes
        cells = _split_cells(_strip_outer_pipes(line))
        cells = [_unescape(c.strip()) for c in cells]
        rows.append(cells)

    return rows


def _is_separator_line(line: str) -> bool:
    """Check if line is a markdown separator like |---|---|."""
    stripped = re.sub(r"[\s|]", "", line)
    if not stripped:
        return False
    return all(c in "-:" for c in stripped)


def _strip_outer_pipes(line: str) -> str:
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        preceding = stripped[:-1]
        backslashes = len(preceding) - len(preceding.rstrip("\\"))
        if backslashes % 2 == 0:
            stripped = preceding
    return stripped


def _split_cells(line: str) -> list[str]:
    """Split delimiters while retaining escapes for one later decoding pass."""
    cells = []
    start = 0
    backslashes = 0
    for index, character in enumerate(line):
        if character == "|" and backslashes % 2 == 0:
            cells.append(line[start:index])
            start = index + 1
        backslashes = backslashes + 1 if character == "\\" else 0
    cells.append(line[start:])
    return cells


def _unescape(text: str) -> str:
    """Decode table-cell escapes once, leaving other Markdown text intact."""
    def decode(match: re.Match[str]) -> str:
        token = match.group()
        if token.startswith("\\"):
            return token[1:]
        if token.startswith("<"):
            return "\n"
        return unescape(token)

    # A single pass keeps &lt;br&gt; literal and &amp;lt; decoded only once.
    return _CELL_ESCAPE.sub(decode, text)
