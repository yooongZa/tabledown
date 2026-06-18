"""Helpers for the Windows CF_HTML clipboard format."""
from __future__ import annotations

import re


CF_HTML_FORMAT_NAME = "HTML Format"


def build_cf_html(fragment: str) -> bytes:
    """Return bytes for the Windows CF_HTML clipboard format."""
    html_prefix = "<html><body><!--StartFragment-->"
    html_suffix = "<!--EndFragment--></body></html>"
    html = f"{html_prefix}{fragment}{html_suffix}"

    header_template = (
        "Version:0.9\r\n"
        "StartHTML:0000000000\r\n"
        "EndHTML:0000000000\r\n"
        "StartFragment:0000000000\r\n"
        "EndFragment:0000000000\r\n"
    )
    header_size = len(header_template.encode("ascii"))
    prefix_size = len(html_prefix.encode("utf-8"))
    fragment_size = len(fragment.encode("utf-8"))
    html_size = len(html.encode("utf-8"))

    start_html = header_size
    end_html = start_html + html_size
    start_fragment = start_html + prefix_size
    end_fragment = start_fragment + fragment_size

    header = (
        "Version:0.9\r\n"
        f"StartHTML:{start_html:010d}\r\n"
        f"EndHTML:{end_html:010d}\r\n"
        f"StartFragment:{start_fragment:010d}\r\n"
        f"EndFragment:{end_fragment:010d}\r\n"
    )
    return header.encode("ascii") + html.encode("utf-8")


def extract_cf_html(data: bytes | str) -> str:
    """Extract the HTML fragment from Windows CF_HTML bytes."""
    raw = _coerce_bytes(data).rstrip(b"\0")
    if not raw:
        return ""

    header = raw[: min(len(raw), 4096)].decode("ascii", errors="ignore")
    start_fragment = _header_int(header, "StartFragment")
    end_fragment = _header_int(header, "EndFragment")
    start_html = _header_int(header, "StartHTML")
    end_html = _header_int(header, "EndHTML")

    start = start_fragment if _valid_offset(raw, start_fragment) else start_html
    end = end_fragment if _valid_offset(raw, end_fragment) else end_html

    if start is None or end is None or start >= end:
        return _ensure_table_wrapper(_decode_html(raw))

    return _ensure_table_wrapper(_decode_html(raw[start:end]))


def _ensure_table_wrapper(html: str) -> str:
    """Wrap a bare table-row fragment in ``<table>`` so parsers find the table.

    Excel/Sheets put the CF_HTML ``StartFragment``/``EndFragment`` markers
    *inside* the ``<table>`` element, so the extracted fragment is the table's
    interior (``<col>``/``<tr>``/``<td>``) with the ``<table>`` open/close tag
    left outside. Downstream table detection keys on ``<table>``, so without the
    wrapper a real Excel table reads as plain text and never converts. Only wrap
    when there are rows but no table tag — never touch non-table HTML.
    """
    lowered = html.lower()
    # Match a real row tag (<tr>, <tr ...>, <tr/>) on a word boundary so
    # <track>/<trail>/<tr-foo> are not mistaken for table rows and wrapped into
    # a bogus <table> (which would then fail to parse).
    if re.search(r"<tr[\s/>]", lowered) and "<table" not in lowered:
        return f"<table>{html}</table>"
    return html


def _coerce_bytes(data: bytes | str) -> bytes:
    if isinstance(data, bytes):
        return data
    return data.encode("utf-8")


def _header_int(header: str, name: str) -> int | None:
    match = re.search(rf"^{name}:(\d+)", header, flags=re.MULTILINE)
    if not match:
        return None
    try:
        return int(match.group(1))
    except ValueError:
        return None


def _valid_offset(raw: bytes, offset: int | None) -> bool:
    return offset is not None and 0 <= offset <= len(raw)


def _decode_html(data: bytes) -> str:
    """Decode CF_HTML bytes to text, matching how macOS reads ready-made HTML.

    CF_HTML is a byte-oriented format spec'd as UTF-8. cp1252 maps all 256 byte
    values, so it is a lossless final fallback. utf-16 is deliberately NOT tried
    by content-sniffing: a valid even-length cp1252/latin fragment can *succeed*
    under utf-16 and silently mojibake (adjacent latin bytes merge into CJK
    glyphs), which then fails table detection and skips a conversion macOS would
    perform — on macOS the HTML arrives as a proper string and never mis-decodes.
    Only honor utf-16 when a byte-order mark actually declares it.
    """
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        try:
            return data.decode("utf-16")
        except UnicodeDecodeError:
            pass
    for encoding in ("utf-8", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")
