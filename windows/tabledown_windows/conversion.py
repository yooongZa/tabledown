"""Platform-independent clipboard conversion decisions for Windows.

Mirrors the macOS ``TabledownApp._converted_clipboard`` and the same clipboard
invariants ("클립보드 변환 불변식 (회귀 금지)" in the repo-root CLAUDE.md): never
strip the HTML slot for web tables, documents, or bare spreadsheet tables, and
apply the strict cell-count check only to html-less text. The win_clipboard
layer preserves any slot not named in ``drop_formats`` — including CF_HTML and
native spreadsheet formats — so leaving HTML out of the drop set keeps it.
"""
from __future__ import annotations

from tablemark.converter.html_to_md import (
    convert_document_tables,
    html_has_content_outside_table,
    html_table_to_markdown,
)
from tablemark.converter.md_to_tsv import is_markdown_table, markdown_table_to_html

from .html_clipboard import CF_HTML_FORMAT_NAME


# The CF_HTML slot is NEVER dropped (invariant 3, unified in 0.2.4): keeping it
# lets a single copy serve both a Markdown editor (reads text) and Excel/Word
# (read the <table> from HTML). Kept here only for read-side symmetry.
WINDOWS_HTML_FORMATS = {CF_HTML_FORMAT_NAME}

# Rendered image/rich snapshots of the table. These ARE dropped so the
# destination cannot paste a stale picture of the table instead of our
# refreshed text + HTML.
WINDOWS_RENDERED_FORMATS = {
    "Bitmap",
    "DeviceIndependentBitmap",
    "DeviceIndependentBitmapV5",
    "EnhancedMetafile",
    "MetafilePict",
    "PNG",
    "Rich Text Format",
}

# Mirrors macOS RENDERED_TABLE_TYPES: only rendered images are dropped; the HTML
# slot is always preserved. (Through 0.2.3 the Windows port also dropped CF_HTML
# here, which broke re-pasting a converted table back into Excel/Word — the same
# loss macOS invariant 1/3 guards against.)
WINDOWS_DROP_FORMATS = WINDOWS_RENDERED_FORMATS


def markdown_paste_block(markdown: str) -> str:
    """Return markdown padded so block parsers see a standalone table."""
    return "\n" + markdown.strip() + "\n"


def converted_clipboard(content: dict) -> dict | None:
    """Return clipboard formats to write, or None when no update is needed."""
    if content.get("generated"):
        return None

    html = content.get("html", "")
    text = content.get("text", "")
    has_html_table = bool(html) and "<table" in html.lower()

    # An accompanying HTML <table> is independent proof of a real table, so skip
    # the strict cell-count check (it only guards html-less text that merely
    # looks like a table). text + html is already multi-format — leave it so the
    # destination picks its own slot (Excel -> table, Markdown editor -> text).
    if text and is_markdown_table(text, strict=not has_html_table):
        if has_html_table:
            return None
        return {
            "text": text,
            "html": markdown_table_to_html(text),
        }

    if has_html_table:
        # A table embedded in a document (headings, paragraphs, lists): render
        # its tables as Markdown in the text slot but KEEP the original HTML so
        # rich editors still paste a real table. Only rendered images are dropped.
        if html_has_content_outside_table(html):
            converted = convert_document_tables(html)
            if not converted.strip() or converted.strip() == text.strip():
                return None
            return {
                "text": converted,
                "drop_formats": WINDOWS_DROP_FORMATS,
            }
        # A bare Excel/Sheets table: Markdown into the text slot, KEEP the HTML
        # <table> so one copy works for every destination. Only rendered images
        # are dropped — never the HTML (invariant 3, unified in 0.2.4).
        markdown = html_table_to_markdown(html)
        if text.strip() == markdown.strip():
            return None
        return {
            "text": markdown_paste_block(markdown),
            "drop_formats": WINDOWS_DROP_FORMATS,
        }

    return None
