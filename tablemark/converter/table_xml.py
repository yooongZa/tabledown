"""Convert between table rows and LLM-friendly XML.

The XML shape is *record-style with structural tags*: one ``<row>`` per data
row, and inside each row a fixed ``<cell name="…">`` per value. A multi-level
column header (a group header spanning sub-columns) is preserved as nested
``<group name="…">`` — the hierarchy the source table drew is kept as real XML
structure instead of being flattened into a string:

    <dataset>
      <row>
        <cell name="직급">부장</cell>
        <cell name="직책">대족장</cell>
        <group name="1분기">
          <cell name="1">동</cell>
          <cell name="2">해</cell>
        </group>
        <group name="2분기">
          <cell name="4">과</cell>
        </group>
      </row>
    </dataset>

Why this shape (see the XML-design analysis behind invariant 5 in CLAUDE.md):

- **Names live in attributes, not tag names.** Column headers are *data* (they
  carry spaces, digits, ``( ) % /``, duplicates, the reserved ``xml`` prefix),
  none of which are legal XML element names. Every real tabular-XML standard
  (OOXML SpreadsheetML, HTML ``<td>``, ODF) uses *fixed* structural tags and
  carries the header as content/attribute. Putting the header in ``name=``
  avoids all the name-mangling the header-as-tag approach forced.
- **The root is ``<dataset>``, never ``<table>``.** ``<table>`` is a real HTML
  element: when this text lands anywhere that renders HTML (a browser, an
  Obsidian preview, a rich-text editor), the HTML parser applies table rules and
  *foster-parents the non-table children out*, emptying the table. ``dataset``/
  ``row``/``cell``/``group`` are not HTML elements, so the tree survives.
- **Horizontal header hierarchy nests; vertical merges do not.** A group header
  over sub-columns is an explicit, author-declared hierarchy → ``<group>``. A
  vertical (rowspan) merge is just "same value repeated", not a hierarchy, so it
  is forward-filled into each ``<row>`` (handled upstream in ``html_to_md``),
  keeping every row a complete, self-contained record.

Parsing is intentionally conservative: ``is_table_xml`` only accepts XML whose
rows contain exactly ``<cell>``/``<group>`` nodes, so a stray config/document
XML on the clipboard is left untouched (same false-positive caution as the
Markdown heuristic). The reverse direction round-trips back to a table model.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

# Root/row tag names that, on their own, are strong evidence of a table. They
# let a single-row table through; otherwise at least two rows are required so a
# one-off nested record (e.g. a config block) is not mistaken for a table.
_KNOWN_ROOT = {
    "dataset", "table", "rows", "records", "data", "rowset", "sheet", "worksheet",
}
_KNOWN_ROW = {"row", "record", "item", "entry", "tr"}

# The fixed structural tags this module emits and recognises.
_GROUP_TAG = "group"
_CELL_TAG = "cell"


def model_to_xml(
    header_rows: list[list[str]],
    data_rows: list[list[str]],
    *,
    root_tag: str = "dataset",
    row_tag: str = "row",
    indent: str = "  ",
) -> str:
    """Render a table model as LLM-friendly record XML.

    ``header_rows`` is one list per header level, column-aligned (merged header
    cells already filled across their span); a single-level header is just a
    one-element list. ``data_rows`` are the (forward-filled) data rows.

    Raises:
        ValueError: if there is no header or no data.
    """
    if not header_rows or not header_rows[0]:
        raise ValueError("헤더가 없습니다")
    if not data_rows:
        raise ValueError("데이터 행이 없습니다")

    ncols = max(len(level) for level in header_rows)
    header_rows = [level + [""] * (ncols - len(level)) for level in header_rows]
    columns = _build_columns(header_rows, list(range(ncols)), 0)

    lines = [f"<{root_tag}>"]
    for row in data_rows:
        lines.append(f"{indent}<{row_tag}>")
        _emit_columns(lines, columns, row, indent, 2)
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


def table_xml_to_model(xml: str) -> tuple[list[list[str]], list[list[str]]]:
    """Parse table XML back into ``(header_rows, data_rows)``.

    Reconstructs the multi-level header from the ``<group>`` nesting so a
    re-conversion round-trips: leaf names land on the deepest header level and
    each group name spans the columns under it.

    Raises:
        ValueError: if the XML does not describe a flat table.
    """
    root = _parse_table_root(xml)
    row_elements = [child for child in root if _is_element(child)]

    tree = _parse_row_tree(row_elements[0])
    depth = max(_tree_depth(tree), 1)
    ncols = _tree_leaf_count(tree)
    header_rows = [[""] * ncols for _ in range(depth)]

    col = 0

    def walk(nodes: list[tuple], level: int) -> None:
        nonlocal col
        for kind, name, children in nodes:
            if kind == "group":
                start = col
                walk(children, level + 1)
                for index in range(start, col):
                    header_rows[level][index] = name
            else:  # leaf — its name belongs on the deepest level
                header_rows[depth - 1][col] = name
                col += 1

    walk(tree, 0)
    data_rows = [_row_values(row) for row in row_elements]
    return header_rows, data_rows


def table_xml_to_rows(xml: str) -> list[list[str]]:
    """Parse table XML into flat rows (first row = combined header).

    Groups are flattened into composite header names ("1분기" over "1" -> the
    single header "1분기 1"), the lossless shape a single-header consumer
    (Markdown) can hold.
    """
    header_rows, data_rows = table_xml_to_model(xml)
    return [_combine_header_levels(header_rows)] + data_rows


def table_xml_to_markdown(xml: str) -> str:
    """Convert table XML to a GFM Markdown table.

    Raises:
        ValueError: if the XML does not describe a flat table.
    """
    return _rows_to_markdown(table_xml_to_rows(xml))


# --- internals: build ---


def _build_columns(
    header_rows: list[list[str]], cols: list[int], level: int
) -> list[tuple]:
    """Recursively turn aligned header levels into a column tree.

    Returns nodes that are either ``("leaf", col_index, name)`` or
    ``("group", name, children)``. Consecutive columns sharing the same label at
    ``level`` form a group; a column with no label at ``level`` (e.g. a key
    column under an empty group cell) is a leaf named by its deepest label.
    """
    nodes: list[tuple] = []
    i = 0
    last_level = level >= len(header_rows) - 1
    while i < len(cols):
        col = cols[i]
        label = header_rows[level][col].strip() if level < len(header_rows) else ""
        if not label or last_level:
            nodes.append(("leaf", col, _leaf_name(header_rows, col)))
            i += 1
        else:
            same = [col]
            j = i + 1
            while j < len(cols) and header_rows[level][cols[j]].strip() == label:
                same.append(cols[j])
                j += 1
            nodes.append(("group", label, _build_columns(header_rows, same, level + 1)))
            i = j
    return nodes


def _leaf_name(header_rows: list[list[str]], col: int) -> str:
    """The deepest non-empty header label for a column (its leaf name)."""
    for level in reversed(header_rows):
        value = level[col].strip()
        if value:
            return value
    return f"col{col + 1}"


def _emit_columns(
    lines: list[str], columns: list[tuple], row: list[str], indent: str, depth: int
) -> None:
    for node in columns:
        if node[0] == "leaf":
            _, col, name = node
            value = row[col] if col < len(row) else ""
            value = value.replace("<br/>", "\n").replace("<br>", "\n")
            lines.append(
                f"{indent * depth}<{_CELL_TAG} name={quoteattr(name)}>"
                f"{escape(value)}</{_CELL_TAG}>"
            )
        else:
            _, name, children = node
            lines.append(f"{indent * depth}<{_GROUP_TAG} name={quoteattr(name)}>")
            _emit_columns(lines, children, row, indent, depth + 1)
            lines.append(f"{indent * depth}</{_GROUP_TAG}>")


# --- internals: parse ---


def _parse_table_root(xml: str) -> ET.Element:
    """Parse and validate that ``xml`` is a flat table; return the root element.

    Conservative on purpose: every row must contain only ``<cell>`` (leaf) or
    ``<group>`` (recursively the same) nodes, so config/document XML is rejected.

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
        nodes = [node for node in row if _is_element(node)]
        if not nodes:
            raise ValueError("row without cells")
        for node in nodes:
            _validate_column_node(node)

    known = (
        _local_name(root.tag).lower() in _KNOWN_ROOT
        or _local_name(row_elements[0].tag).lower() in _KNOWN_ROW
    )
    if len(row_elements) < 2 and not known:
        raise ValueError("ambiguous single-row xml")

    return root


def _validate_column_node(element: ET.Element) -> None:
    """A column node is a leaf ``<cell>`` or a non-empty ``<group>`` of nodes."""
    name = _local_name(element.tag).lower()
    if name == _GROUP_TAG:
        children = [child for child in element if _is_element(child)]
        if not children:
            raise ValueError("empty group")
        for child in children:
            _validate_column_node(child)
    elif name == _CELL_TAG:
        if any(_is_element(child) for child in element):
            raise ValueError("nested cell")
    else:
        raise ValueError(f"unexpected node <{name}>")


def _parse_row_tree(row: ET.Element) -> list[tuple]:
    """Turn one ``<row>`` into a column tree of ``(kind, name, children)``."""
    nodes: list[tuple] = []
    for child in row:
        if not _is_element(child):
            continue
        name = child.get("name") or _local_name(child.tag)
        if _local_name(child.tag).lower() == _GROUP_TAG:
            nodes.append(("group", name, _parse_row_tree(child)))
        else:
            nodes.append(("leaf", name, []))
    return nodes


def _row_values(row: ET.Element) -> list[str]:
    """Depth-first leaf-cell text values for one ``<row>``."""
    values: list[str] = []
    for child in row:
        if not _is_element(child):
            continue
        if _local_name(child.tag).lower() == _GROUP_TAG:
            values.extend(_row_values(child))
        else:
            values.append(child.text or "")
    return values


def _tree_depth(tree: list[tuple]) -> int:
    depth = 1
    for kind, _, children in tree:
        if kind == "group":
            depth = max(depth, 1 + _tree_depth(children))
    return depth


def _tree_leaf_count(tree: list[tuple]) -> int:
    count = 0
    for kind, _, children in tree:
        count += _tree_leaf_count(children) if kind == "group" else 1
    return count


def _combine_header_levels(header_rows: list[list[str]]) -> list[str]:
    """Join multi-level header labels per column ("매출" / "1분기" -> "매출 1분기")."""
    if not header_rows:
        return []
    ncols = max(len(level) for level in header_rows)
    combined = []
    for col in range(ncols):
        labels: list[str] = []
        for level in header_rows:
            value = (level[col] if col < len(level) else "").strip()
            if value and (not labels or labels[-1] != value):
                labels.append(value)
        combined.append(" ".join(labels))
    return combined


# --- internals: markdown + shared ---


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
