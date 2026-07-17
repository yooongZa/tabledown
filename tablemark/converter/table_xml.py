"""Convert between a table model and LLM-friendly **nested** XML (format v2).

The v2 shape preserves a two-dimensional header hierarchy (vertical *and*
horizontal groups) as real XML **nesting** instead of flattening it into a
string. A parent group is a parent node; a sub-item is a child node; a label is
an attribute; a cell value is leaf text:

    <표>
      <직급그룹 이름="부장">
        <행 직책="대족장">
          <열그룹 이름="1분기"><열 n="1">동</열><열 n="2">해</열><열 n="3">물</열></열그룹>
          <열그룹 이름="2분기"><열 n="4">과</열><열 n="5">백</열><열 n="6">두</열></열그룹>
        </행>
        <행 직책="족장"> … </행>
      </직급그룹>
      <직급그룹 이름="차장"> … </직급그룹>
    </표>

Tag / attribute mapping (spec §2):

===================  ====================  ==============  ===========
table element        tag                   attribute       text
===================  ====================  ==============  ===========
root                 ``<표>``              —               —
vertical group       ``<{header}그룹>``     ``이름="값"``    —
row (deepest key)    ``<행>``              ``{header}="값"`` —
horizontal group     ``<열그룹>``           ``이름="값"``    —
horizontal leaf      ``<열>``              ``n="헤더값"``   cell value
===================  ====================  ==============  ===========

Why this shape (see the XML-design analysis behind invariant 5 in CLAUDE.md):

- **The root is ``<표>`` (Korean), never ``<table>``.** ``<table>`` is a real
  HTML element: anywhere this text lands that renders HTML (a browser, an
  Obsidian preview, a rich-text editor) the HTML parser applies table rules and
  *foster-parents the non-table children out*, emptying it. ``표``/``행``/``열``/
  ``열그룹`` are not HTML elements, so the tree survives — the same reason v1 used
  ``<dataset>``.
- **Horizontal labels live in attributes; the horizontal tags are fixed.**
  ``<열그룹 이름>``/``<열 n>`` keep their labels in *attribute values*, which can
  hold spaces, digits, ``( ) % /``, duplicates, the reserved ``xml`` prefix —
  none of which are legal element names.
- **Vertical keys nest by header tag.** The left key columns become nested
  ``<{header}그룹 이름="값">`` (all but the last) and ``<행 {header}="값">`` (the
  last). A vertical key header is therefore used as a *tag name*, so it must be a
  valid XML Name; key columns whose header is not a valid Name are demoted to
  data (``_valid_xml_name`` / spec §3-1) — no mangling, lossless round-trip.

The biggest change from v1 is **vertical grouping**: v1 repeated the left key
columns on every row (flat, self-contained records); v2 nests them, so a single
row is no longer self-contained (spec §7) — the hierarchy is preserved instead.

Parsing is intentionally conservative (spec §6): ``is_table_xml`` only accepts
XML whose rows are ``<행>`` holding ``<열>``/``<열그룹>`` (or the analogous v1/
generic shape), so a stray config/document XML on the clipboard is left
untouched (same false-positive caution as the Markdown heuristic). The Markdown
path (``table_xml_to_rows``/``table_xml_to_markdown``) flattens the nested XML
back to combined-header rows, the lossless shape a single-header consumer can
hold.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

from .formula_export import validate_xml_text

# Root/row tag names that, on their own, are strong evidence of a table. They
# let a single-row table through; otherwise at least two rows are required so a
# one-off nested record (e.g. a config block) is not mistaken for a table.
_KNOWN_ROOT = {
    "표", "dataset", "table", "rows", "records", "data", "rowset", "sheet", "worksheet",
}
_KNOWN_ROW = {"행", "row", "record", "item", "entry", "tr"}

# The fixed structural tags v2 emits. ``열그룹`` is the horizontal group, ``열``
# the horizontal leaf (cell), ``행`` the row. A vertical group tag is
# ``"{header}" + _GROUP_SUFFIX`` (e.g. ``직급그룹``) — not a fixed name.
_ROOT_TAG = "표"
_ROW_TAG = "행"
_COL_GROUP_TAG = "열그룹"
_COL_TAG = "열"
_GROUP_SUFFIX = "그룹"
_NAME_ATTR = "이름"  # group label attribute (both vertical groups and 열그룹)
_LEAF_ATTR = "n"     # 열 leaf header attribute

# v1 tags still recognised on input so already-copied v1 XML keeps round-tripping.
_V1_GROUP_TAG = "group"
_V1_CELL_TAG = "cell"


def model_to_xml(
    header_rows: list[list[str]],
    data_rows: list[list[str]],
    *,
    indent: str = "  ",
) -> str:
    """Render a table model as nested v2 XML (spec §3).

    ``header_rows`` is one list per header level, column-aligned (merged header
    cells already filled across their span); a single-level header is just a
    one-element list. ``data_rows`` are the (forward-filled) data rows, including
    the left vertical-key columns.

    Vertical key columns are the leftmost consecutive top-level *leaves* before
    the first horizontal ``group`` whose header is a valid XML Name. With no
    horizontal group there are zero vertical keys (a simple table, spec §4-2).

    Raises:
        ValueError: if there is no header or no data.
    """
    if not header_rows or not header_rows[0]:
        raise ValueError("헤더가 없습니다")
    if not data_rows:
        raise ValueError("데이터 행이 없습니다")

    # quoteattr()/escape() encode markup characters but do not reject control
    # code points forbidden by XML 1.0. Validate every source field before
    # emitting anything so callers never receive malformed XML.
    for row in [*header_rows, *data_rows]:
        for value in row:
            validate_xml_text(value)

    ncols = max(len(level) for level in header_rows)
    header_rows = [level + [""] * (ncols - len(level)) for level in header_rows]

    key_headers, data_cols = _split_vertical_keys(header_rows, ncols)
    # Build the horizontal column tree over the *data* columns only, so a key
    # column (same label on every level) is not mistaken for a 1-column group.
    data_columns = _build_columns(header_rows, data_cols, 0)

    lines = [f"<{_ROOT_TAG}>"]
    _emit_rows(lines, key_headers, data_columns, data_rows, 0, len(data_rows), 0, indent, 1)
    lines.append(f"</{_ROOT_TAG}>")
    return "\n".join(lines)


def is_table_xml(text: str) -> bool:
    """Heuristic: does ``text`` parse as XML that describes a table? (spec §6)"""
    try:
        _parse_table_root(text)
    except ValueError:
        return False
    return True


def table_xml_to_model(xml: str) -> tuple[list[list[str]], list[list[str]]]:
    """Parse v2 (or v1) table XML back into ``(header_rows, data_rows)`` (spec §5).

    Reconstructs both the vertical-key column headers (from the ``<{h}그룹>``/
    ``<행 {h}=>`` nesting) and the horizontal multi-level header (from the
    ``<열그룹>``/``<열>`` nesting), and forward-fills the vertical keys into each
    row, so a re-conversion with ``model_to_xml`` round-trips.

    Raises:
        ValueError: if the XML does not describe a table.
    """
    root = _parse_table_root(xml)

    # Walk the vertical nesting: accumulate (key_values, row_element) per row,
    # carrying each enclosing group's (header, value) down to its rows.
    rows: list[tuple[list[tuple[str, str]], ET.Element]] = []

    def walk(node: ET.Element, keys: list[tuple[str, str]]) -> None:
        tag = _local_name(node.tag)
        if _is_row_tag(node):
            row_keys = list(keys)
            for attr, value in node.attrib.items():
                row_keys.append((_local_name(attr), value))
            rows.append((row_keys, node))
            return
        # a vertical group <{header}그룹 이름="값"> — strip the suffix for header
        header = tag[: -len(_GROUP_SUFFIX)] if tag.endswith(_GROUP_SUFFIX) else tag
        value = node.get(_NAME_ATTR, "")
        child_keys = keys + [(header, value)]
        for child in node:
            if _is_element(child):
                walk(child, child_keys)

    for child in root:
        if _is_element(child):
            walk(child, [])

    # Vertical key headers, in column order, from the first row.
    key_headers = [header for header, _ in rows[0][0]] if rows else []
    nkeys = len(key_headers)

    # Horizontal column tree from the first row's <열>/<열그룹> children.
    tree = _parse_row_tree(rows[0][1])
    h_depth = max(_tree_depth(tree), 1)
    h_ncols = _tree_leaf_count(tree)

    ncols = nkeys + h_ncols
    depth = max(h_depth, 1)
    # Key headers repeat on every header level (a vertical merge), so they fill
    # the whole header column; horizontal labels land per the tree's depth.
    header_rows = [[""] * ncols for _ in range(depth)]
    for level in range(depth):
        for index, header in enumerate(key_headers):
            header_rows[level][index] = header

    col = nkeys

    def fill_header(nodes: list[tuple], level: int) -> None:
        nonlocal col
        for kind, name, children in nodes:
            if kind == "group":
                start = col
                fill_header(children, level + 1)
                for index in range(start, col):
                    header_rows[level][index] = name
            else:  # leaf — its name belongs on the deepest level
                header_rows[depth - 1][col] = name
                col += 1

    fill_header(tree, 0)

    data_rows = []
    for row_keys, row_element in rows:
        values = [value for _, value in row_keys] + _row_values(row_element)
        data_rows.append(values)
    return header_rows, data_rows


def table_xml_to_rows(xml: str) -> list[list[str]]:
    """Parse table XML into flat rows (first row = combined header).

    Groups are flattened into composite header names ("1분기" over "1" -> the
    single header "1분기 1"), the lossless shape a single-header consumer
    (Markdown) can hold (spec §7).
    """
    header_rows, data_rows = table_xml_to_model(xml)
    return [_combine_header_levels(header_rows)] + data_rows


def table_xml_to_markdown(xml: str) -> str:
    """Convert table XML to a GFM Markdown table.

    Raises:
        ValueError: if the XML does not describe a table.
    """
    return _rows_to_markdown(table_xml_to_rows(xml))


# --- internals: build ---


def _split_vertical_keys(
    header_rows: list[list[str]], ncols: int
) -> tuple[list[str], list[int]]:
    """Split columns into (vertical key headers, data column indices) (spec §3-1).

    A vertical key column is a leading column that is a *leaf* (it carries no
    horizontal grouping) — every level above the deepest is either blank or
    repeats the deepest label. This covers both real Excel shapes: the leaf
    sitting at the bottom with empty cells above (``''`` over ``직급``) and a
    full-height header merge (``직급`` over ``직급``). A horizontal group, by
    contrast, has an upper label fanning out into *different* sub-labels below.
    Keys must also be valid XML Names (their header becomes a tag); the first
    column that is not such a leaf, or not a valid Name, ends the key region.

    Per spec §3-1 the keys only apply when the table actually has a horizontal
    ``group``: with no group at all (a simple/flat table) there are zero keys and
    every column is data (spec §4-2). ``ncols`` is the (padded) column count.
    """
    all_cols = list(range(ncols))

    # Candidate leading keys: leaf columns (no horizontal grouping), valid Names.
    candidate = 0
    for col in all_cols:
        name = _leaf_name(header_rows, col)
        if not _column_is_leaf(header_rows, col) or not _valid_xml_name(name):
            break
        candidate += 1

    # Only treat them as keys if the remaining columns contain a real horizontal
    # group; otherwise it is a simple table → no vertical keys.
    data_cols = all_cols[candidate:]
    if not _has_group(header_rows, data_cols):
        return [], all_cols

    key_headers = [_leaf_name(header_rows, col) for col in all_cols[:candidate]]
    return key_headers, data_cols


def _column_is_leaf(header_rows: list[list[str]], col: int) -> bool:
    """True if column ``col`` carries no horizontal grouping (a key-eligible leaf).

    Every header level above the deepest must be blank or repeat the deepest
    label; a level with a *different* non-empty label means the column is under a
    horizontal group (e.g. ``매출`` over ``1분기``) and is not a key leaf.
    """
    leaf = _leaf_name(header_rows, col)
    for level in header_rows[:-1]:
        label = (level[col] if col < len(level) else "").strip()
        if label and label != leaf:
            return False
    return True


def _has_group(header_rows: list[list[str]], cols: list[int]) -> bool:
    """True if ``cols`` form any multi-column horizontal group (a real header tree)."""
    return any(node[0] == "group" for node in _build_columns(header_rows, cols, 0))


def _emit_rows(
    lines: list[str],
    key_headers: list[str],
    data_columns: list[tuple],
    data_rows: list[list[str]],
    start: int,
    end: int,
    key_level: int,
    indent: str,
    depth: int,
) -> None:
    """Emit rows[start:end] as nested vertical groups then ``<행>`` elements.

    ``key_level`` indexes ``key_headers``: levels before the last become nested
    ``<{header}그룹 이름="값">`` (consecutive equal values grouped into runs); the
    last key header (or none) terminates at ``<행>``.
    """
    nkeys = len(key_headers)
    pad = indent * depth

    if key_level >= nkeys - 1:
        # Deepest key (or no keys): emit a <행> per row.
        last_attr = key_headers[key_level] if key_level < nkeys else None
        for r in range(start, end):
            row = data_rows[r]
            if last_attr is not None:
                value = row[key_level] if key_level < len(row) else ""
                lines.append(f"{pad}<{_ROW_TAG} {last_attr}={quoteattr(value)}>")
            else:
                lines.append(f"{pad}<{_ROW_TAG}>")
            _emit_columns(lines, data_columns, row, indent, depth + 1)
            lines.append(f"{pad}</{_ROW_TAG}>")
        return

    # An intermediate key level: group consecutive equal values into a run, each
    # a <{header}그룹 이름="값"> that recurses on the next key level.
    header = key_headers[key_level]
    r = start
    while r < end:
        value = data_rows[r][key_level] if key_level < len(data_rows[r]) else ""
        run_end = r + 1
        while run_end < end:
            nxt = data_rows[run_end][key_level] if key_level < len(data_rows[run_end]) else ""
            if nxt != value:
                break
            run_end += 1
        lines.append(f"{pad}<{header}{_GROUP_SUFFIX} {_NAME_ATTR}={quoteattr(value)}>")
        _emit_rows(
            lines, key_headers, data_columns, data_rows,
            r, run_end, key_level + 1, indent, depth + 1,
        )
        lines.append(f"{pad}</{header}{_GROUP_SUFFIX}>")
        r = run_end


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
    """Emit the horizontal column tree for one row: ``<열그룹>`` / ``<열>``."""
    for node in columns:
        if node[0] == "leaf":
            _, col, name = node
            value = row[col] if col < len(row) else ""
            lines.append(
                f"{indent * depth}<{_COL_TAG} {_LEAF_ATTR}={quoteattr(name)}>"
                f"{escape(value)}</{_COL_TAG}>"
            )
        else:
            _, name, children = node
            lines.append(
                f"{indent * depth}<{_COL_GROUP_TAG} {_NAME_ATTR}={quoteattr(name)}>"
            )
            _emit_columns(lines, children, row, indent, depth + 1)
            lines.append(f"{indent * depth}</{_COL_GROUP_TAG}>")


# --- internals: parse ---


def _parse_table_root(xml: str) -> ET.Element:
    """Parse and validate that ``xml`` is a table; return the root element (spec §6).

    Conservative on purpose: every row is ``<행>`` (or a known/generic row tag)
    whose children are only ``<열>``/``<열그룹>`` (or v1 ``<cell>``/``<group>``),
    and the vertical nesting may only contain ``*그룹`` group wrappers and rows —
    so config/document XML is rejected.

    Raises:
        ValueError: with a reason when it is not a table.
    """
    stripped = xml.strip()
    if not stripped.startswith("<"):
        raise ValueError("not xml")
    try:
        root = ET.fromstring(stripped)
    except ET.ParseError as exc:
        # Keep the ParseError chained for local debugging (``from exc``) but keep
        # it OUT of the message text: the parser detail can echo clipboard markup,
        # and this message may reach the user-shareable diagnostics log.
        raise ValueError("xml parse error") from exc

    rows: list[ET.Element] = []
    _collect_rows(root, rows)
    if not rows:
        raise ValueError("no rows")

    # All rows must share one tag (a single dominant row shape).
    if len({_local_name(row.tag) for row in rows}) != 1:
        raise ValueError("inconsistent row tags")

    for row in rows:
        nodes = [node for node in row if _is_element(node)]
        if not nodes:
            raise ValueError("row without cells")
        for node in nodes:
            _validate_column_node(node)

    known = (
        _local_name(root.tag).lower() in _KNOWN_ROOT
        or _local_name(rows[0].tag).lower() in _KNOWN_ROW
    )
    if len(rows) < 2 and not known:
        raise ValueError("ambiguous single-row xml")

    return root


def _collect_rows(node: ET.Element, rows: list[ET.Element]) -> None:
    """Collect all row elements, descending only through ``*그룹`` group wrappers.

    A direct child of the root (or of a vertical group) must be either a row or a
    vertical-group wrapper whose tag ends in ``그룹``; anything else makes this
    not a table. This keeps the guard tight: a config block whose elements are
    neither rows nor ``*그룹`` groups is rejected.
    """
    for child in node:
        if not _is_element(child):
            continue
        if _is_row_tag(child):
            rows.append(child)
        elif _local_name(child.tag).endswith(_GROUP_SUFFIX) and len(_local_name(child.tag)) > len(_GROUP_SUFFIX):
            # a vertical group wrapper (e.g. 직급그룹); must contain rows/groups
            grand = [g for g in child if _is_element(g)]
            if not grand:
                raise ValueError("empty vertical group")
            _collect_rows(child, rows)
        else:
            # The tag name is clipboard-derived (a header value in this format),
            # so keep it out of the message — this can reach the shared log.
            raise ValueError("unexpected node under table")


def _is_row_tag(element: ET.Element) -> bool:
    """True if ``element`` is a row: the v2 ``행`` tag or a known generic row tag."""
    name = _local_name(element.tag).lower()
    return name == _ROW_TAG or name in _KNOWN_ROW


def _validate_column_node(element: ET.Element) -> None:
    """A column node is a leaf (``열``/``cell``) or a non-empty group (``열그룹``/``group``)."""
    name = _local_name(element.tag).lower()
    if name in (_COL_GROUP_TAG, _V1_GROUP_TAG):
        children = [child for child in element if _is_element(child)]
        if not children:
            raise ValueError("empty group")
        for child in children:
            _validate_column_node(child)
    elif name in (_COL_TAG, _V1_CELL_TAG):
        if any(_is_element(child) for child in element):
            raise ValueError("nested cell")
    else:
        # ``name`` is a clipboard-derived tag (a header value); keep it out of
        # the message so it can't leak into the shared diagnostics log.
        raise ValueError("unexpected node")


def _parse_row_tree(row: ET.Element) -> list[tuple]:
    """Turn one row into a horizontal column tree of ``(kind, name, children)``."""
    nodes: list[tuple] = []
    for child in row:
        if not _is_element(child):
            continue
        name = _local_name(child.tag)
        if name in (_COL_GROUP_TAG, _V1_GROUP_TAG):
            label = child.get(_NAME_ATTR) or child.get("name") or name
            nodes.append(("group", label, _parse_row_tree(child)))
        else:  # 열 / cell leaf
            label = child.get(_LEAF_ATTR) or child.get("name") or name
            nodes.append(("leaf", label, []))
    return nodes


def _row_values(row: ET.Element) -> list[str]:
    """Depth-first leaf-cell text values for one row (horizontal cells only)."""
    values: list[str] = []
    for child in row:
        if not _is_element(child):
            continue
        name = _local_name(child.tag)
        if name in (_COL_GROUP_TAG, _V1_GROUP_TAG):
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


# --- internals: XML Name validation ---

# Start chars: letters, underscore, and the CJK / general Unicode ranges the XML
# 1.0 NameStartChar production allows (so Korean/CJK headers are valid tags).
def _is_name_start_char(ch: str) -> bool:
    cp = ord(ch)
    return (
        ch == "_"
        or ("A" <= ch <= "Z")
        or ("a" <= ch <= "z")
        or 0xC0 <= cp <= 0xD6
        or 0xD8 <= cp <= 0xF6
        or 0xF8 <= cp <= 0x2FF
        or 0x370 <= cp <= 0x37D
        or 0x37F <= cp <= 0x1FFF
        or 0x200C <= cp <= 0x200D
        or 0x2070 <= cp <= 0x218F
        or 0x2C00 <= cp <= 0x2FEF
        or 0x3001 <= cp <= 0xD7FF
        or 0xF900 <= cp <= 0xFDCF
        or 0xFDF0 <= cp <= 0xFFFD
        or 0x10000 <= cp <= 0xEFFFF
    )


def _is_name_char(ch: str) -> bool:
    cp = ord(ch)
    return (
        _is_name_start_char(ch)
        or ch in "-."
        or ("0" <= ch <= "9")
        or cp == 0xB7
        or 0x0300 <= cp <= 0x036F
        or 0x203F <= cp <= 0x2040
    )


def _valid_xml_name(name: str) -> bool:
    """True if ``name`` is a usable XML element Name (spec §3-1 vertical-key rule).

    Enforces the XML 1.0 Name production (start-char + body-chars) and rejects
    the reserved ``xml`` prefix. Used to decide whether a left key column's
    header may become a vertical-group / row tag; if not, the column is demoted
    to data so nothing is mangled and the round-trip stays lossless.
    """
    if not name:
        return False
    if not _is_name_start_char(name[0]):
        return False
    if not all(_is_name_char(ch) for ch in name[1:]):
        return False
    # The "xml" prefix (any case) is reserved by the spec.
    if name[:3].lower() == "xml":
        return False
    return True


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
