"""NSPasteboard wrapper for reading HTML + text from macOS clipboard.

Why not pyperclip? Because Excel puts both HTML and TSV on the clipboard,
and pyperclip only sees the text portion. We need NSPasteboard to read HTML.
"""
from AppKit import NSPasteboard, NSPasteboardTypeHTML, NSPasteboardTypeString

LEGACY_HTML_TYPE = "Apple HTML pasteboard type"
LEGACY_STRING_TYPE = "NSStringPboardType"


def clipboard_change_count() -> int:
    """Return the current clipboard change counter."""
    return int(NSPasteboard.generalPasteboard().changeCount())


def read_clipboard() -> dict:
    """Read both HTML and plain text from clipboard.

    Returns:
        dict with optional 'html' and 'text' keys.
    """
    pb = NSPasteboard.generalPasteboard()
    result = {}

    html = pb.stringForType_(NSPasteboardTypeHTML)
    if html:
        result["html"] = str(html)

    text = pb.stringForType_(NSPasteboardTypeString)
    if text:
        result["text"] = str(text)

    return result


def write_clipboard(text: str | None = None, html: str | None = None) -> None:
    """Update text/html formats while preserving existing clipboard formats."""
    pb = NSPasteboard.generalPasteboard()
    existing_data = {}
    for pb_type in pb.types() or []:
        pb_type = str(pb_type)
        data = pb.dataForType_(pb_type)
        if data is not None:
            existing_data[pb_type] = data

    string_types = {str(NSPasteboardTypeString), LEGACY_STRING_TYPE}
    html_types = {str(NSPasteboardTypeHTML), LEGACY_HTML_TYPE}
    replaced_types = set()
    if text is not None:
        replaced_types.update(string_types)
    if html is not None:
        replaced_types.update(html_types)

    declared_types = []
    for pb_type in existing_data:
        if pb_type not in declared_types:
            declared_types.append(pb_type)
    for pb_type in [str(NSPasteboardTypeString), LEGACY_STRING_TYPE]:
        if text is not None and pb_type not in declared_types:
            declared_types.append(pb_type)
    for pb_type in [str(NSPasteboardTypeHTML), LEGACY_HTML_TYPE]:
        if html is not None and pb_type not in declared_types:
            declared_types.append(pb_type)

    pb.clearContents()
    pb.declareTypes_owner_(declared_types, None)

    for pb_type, data in existing_data.items():
        if pb_type not in replaced_types:
            pb.setData_forType_(data, pb_type)
    if text is not None:
        for pb_type in string_types:
            pb.setString_forType_(text, pb_type)
    if html is not None:
        for pb_type in html_types:
            pb.setString_forType_(html, pb_type)
