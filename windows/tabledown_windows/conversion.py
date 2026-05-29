"""Platform-independent clipboard conversion decisions for Windows."""
from __future__ import annotations

from tablemark.converter.html_to_md import html_table_to_markdown
from tablemark.converter.md_to_tsv import is_markdown_table, markdown_table_to_html

from .html_clipboard import CF_HTML_FORMAT_NAME


WINDOWS_HTML_FORMATS = {CF_HTML_FORMAT_NAME}
WINDOWS_RENDERED_FORMATS = {
    "Bitmap",
    "DeviceIndependentBitmap",
    "DeviceIndependentBitmapV5",
    "EnhancedMetafile",
    "MetafilePict",
    "PNG",
    "Rich Text Format",
}
WINDOWS_DROP_FORMATS = WINDOWS_HTML_FORMATS | WINDOWS_RENDERED_FORMATS


def markdown_paste_block(markdown: str) -> str:
    """Return markdown padded so block parsers see a standalone table."""
    return "\n" + markdown.strip() + "\n"


def converted_clipboard(content: dict) -> dict | None:
    """Return clipboard formats to write, or None when no update is needed."""
    if content.get("generated"):
        return None

    html = content.get("html", "")
    text = content.get("text", "")

    if text and is_markdown_table(text):
        if html and "<table" in html.lower():
            return None
        return {
            "text": text,
            "html": markdown_table_to_html(text),
        }

    if html and "<table" in html.lower():
        markdown = html_table_to_markdown(html)
        if text.strip() == markdown.strip():
            return None
        return {
            "text": markdown_paste_block(markdown),
            "drop_formats": WINDOWS_DROP_FORMATS,
        }

    return None
