"""Convert between table rows and LLM-friendly XML.

The XML shape is *record-style*: one ``<row>`` per data row, and inside each
row the column header becomes the tag for its value:

    <table>
      <row>
        <Name>Alice</Name>
        <Score>95</Score>
      </row>
      <row>
        <Name>Bob</Name>
        <Score>82</Score>
      </row>
    </table>

This is the most LLM-legible XML for tabular data — every value carries its
column name inline, so a model reading the prompt sees ``<Score>95</Score>``
instead of an opaque cell — and it round-trips back to a table.

Header → tag uses XML name rules (Korean/ASCII letters stay verbatim; spaces and
symbols collapse to ``_``). When the tag had to differ from the original header,
the original is preserved in a ``header`` attribute so the reverse direction is
lossless:

    <Q1_2024 header="Q1 2024">12</Q1_2024>

Parsing is intentionally conservative: ``is_table_xml`` only accepts XML that
clearly describes a flat table, so a stray config/document XML on the clipboard
is left untouched (same false-positive caution as the Markdown heuristic).
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

# Root/row tag names that, on their own, are strong evidence of a table. They
# let a single-row table through; otherwise at least two rows are required so a
# one-off nested record (e.g. a config block) is not mistaken for a table.
_KNOWN_ROOT = {
    "table", "rows", "records", "data", "rowset", "dataset", "sheet", "worksheet",
}
_KNOWN_ROW = {"row", "record", "item", "entry", "tr"}

# Anything that is not an XML NameChar (Unicode letters/digits, ``_``, ``-``,
# ``.``) becomes a separator. ``\w`` is Unicode-aware so Hangul/CJK survive.
_INVALID_NAME_CHARS = re.compile(r"[^\w.\-]", re.UNICODE)


def rows_to_xml(
    rows: list[list[str]],
    *,
    root_tag: str = "table",
    row_tag: str = "row",
    indent: str = "  ",
) -> str:
    """Render rows (first row = header) as LLM-friendly record XML.

    Raises:
        ValueError: if ``rows`` is empty.
    """
    if not rows:
        raise ValueError("빈 표입니다")

    headers = rows[0]
    data = rows[1:]

    specs = _column_specs(headers)

    lines = [f"<{root_tag}>"]
    for row in data:
        lines.append(f"{indent}<{row_tag}>")
        for index, (tag, original) in enumerate(specs):
            value = row[index] if index < len(row) else ""
            value = value.replace("<br/>", "\n").replace("<br>", "\n")
            attr = f" header={quoteattr(original)}" if original is not None else ""
            lines.append(f"{indent * 2}<{tag}{attr}>{escape(value)}</{tag}>")
        lines.append(f"{indent}</{row_tag}>")
    lines.append(f"</{root_tag}>")
    return "\n".join(lines)


def is_table_xml(text: str) -> bool:
    """Heuristic: does ``text`` parse as XML that describes a flat table?"""
    try:
        _parse_table_root(text)
    except ValueError:
        return False
    return True


def table_xml_to_rows(xml: str) -> list[list[str]]:
    """Parse table XML into rows of cell values (first row = header).

    Raises:
        ValueError: if the XML does not describe a flat table.
    """
    root = _parse_table_root(xml)
    row_elements = [child for child in root if _is_element(child)]

    headers = [
        cell.get("header") or _local_name(cell.tag)
        for cell in row_elements[0]
        if _is_element(cell)
    ]

    rows = [headers]
    for row in row_elements:
        rows.append([(cell.text or "") for cell in row if _is_element(cell)])
    return rows


def table_xml_to_markdown(xml: str) -> str:
    """Convert table XML to a GFM Markdown table.

    Raises:
        ValueError: if the XML does not describe a flat table.
    """
    return _rows_to_markdown(table_xml_to_rows(xml))


# --- internals ---


def _column_specs(headers: list[str]) -> list[tuple[str, str | None]]:
    """Map headers to ``(tag, original_or_None)``; original kept when tag differs."""
    specs: list[tuple[str, str | None]] = []
    used: set[str] = set()
    for header in headers:
        base = _sanitize_tag(header)
        tag = base
        suffix = 1
        while tag in used:
            suffix += 1
            tag = f"{base}_{suffix}"
        used.add(tag)
        original = None if tag == header else header
        specs.append((tag, original))
    return specs


def _sanitize_tag(name: str) -> str:
    """Coerce a header into a valid XML element name."""
    cleaned = _INVALID_NAME_CHARS.sub("_", name.strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned:
        return "column"
    # XML names may not start with a digit, '-', '.', or the reserved "xml".
    if cleaned[0].isdigit() or cleaned[0] in "-.":
        cleaned = "_" + cleaned
    if cleaned[:3].lower() == "xml":
        cleaned = "_" + cleaned
    return cleaned


def _parse_table_root(xml: str) -> ET.Element:
    """Parse and validate that ``xml`` is a flat table; return the root element.

    Raises:
        ValueError: with a reason when it is not a flat table.
    """
    stripped = xml.strip()
    if not stripped.startswith("<"):
        raise ValueError("not xml")
    try:
        root = ET.fromstring(stripped)
    except ET.ParseError as exc:
        raise ValueError(f"xml parse error: {exc}") from exc

    row_elements = [child for child in root if _is_element(child)]
    if not row_elements:
        raise ValueError("no rows")

    if len({_local_name(row.tag) for row in row_elements}) != 1:
        raise ValueError("inconsistent row tags")

    for row in row_elements:
        cells = [cell for cell in row if _is_element(cell)]
        if not cells:
            raise ValueError("row without cells")
        for cell in cells:
            if any(_is_element(grandchild) for grandchild in cell):
                raise ValueError("nested cell")

    known = (
        _local_name(root.tag).lower() in _KNOWN_ROOT
        or _local_name(row_elements[0].tag).lower() in _KNOWN_ROW
    )
    if len(row_elements) < 2 and not known:
        raise ValueError("ambiguous single-row xml")

    return root


def _rows_to_markdown(rows: list[list[str]]) -> str:
    """Render rows (first row = header) as a GFM Markdown table."""
    if not rows:
        raise ValueError("빈 표입니다")

    max_cols = max(len(row) for row in rows)
    rows = [row + [""] * (max_cols - len(row)) for row in rows]

    header = rows[0]
    body = rows[1:]
    lines = [
        "| " + " | ".join(_markdown_cell(cell) for cell in header) + " |",
        "| " + " | ".join(["---"] * max_cols) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(_markdown_cell(cell) for cell in row) + " |")
    return "\n".join(lines)


def _markdown_cell(value: str) -> str:
    """Escape pipes and keep in-cell line breaks as <br> (GFM), never empty."""
    value = value.replace("\r\n", "\n").replace("\r", "\n").replace("|", "\\|")
    value = "<br>".join(" ".join(line.split()) for line in value.split("\n"))
    return value if value.strip() else " "


def _local_name(tag: str) -> str:
    """Drop an XML namespace prefix like ``{ns}row`` -> ``row``."""
    if isinstance(tag, str) and tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def _is_element(node) -> bool:
    """True for real elements; ElementTree comments/PIs have a callable tag."""
    return isinstance(node.tag, str)
