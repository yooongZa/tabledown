"""Expose generated Markdown to Chromium/Electron paste events, alongside HTML.

Chromium's DataTransfer custom-data format is a little-endian base::Pickle:
payload size, entry count, then UTF-16 strings with 32-bit lengths/alignment.
See ui/base/clipboard/custom_data_helper.cc and base/pickle.cc in Chromium.
This is unrelated to Python pickle; no executable objects are deserialized.
"""
import struct

MAC_WEB_CUSTOM_TYPE = "org.chromium.web-custom-data"
WINDOWS_WEB_CUSTOM_TYPE = "Chromium Web Custom MIME Data Format"
_MAX_BYTES = 16 * 1024 * 1024


def with_markdown(markdown: str, existing: bytes | None = None) -> bytes | None:
    """Merge text/markdown, retaining other custom formats.

    Unknown/malformed or oversized data is left untouched by returning None.
    Call before clearing the clipboard, under the writer's generation guard.
    """
    entries = {}
    if existing is not None:
        if len(existing) < 8 or len(existing) > _MAX_BYTES:
            return None
        size, count = struct.unpack_from("<II", existing)
        end = 4 + size
        if end > len(existing) or size < 4 or count > (size - 4) // 8:
            return None
        offset = 8
        try:
            for _ in range(count):
                pair = []
                for _ in range(2):
                    if offset + 4 > end:
                        return None
                    length = struct.unpack_from("<I", existing, offset)[0] * 2
                    offset += 4
                    if offset + length > end:
                        return None
                    pair.append(existing[offset:offset + length].decode("utf-16-le"))
                    offset += (length + 3) & ~3
                key, value = pair
                if key in entries:
                    return None
                entries[key] = value
        except UnicodeError:
            return None
        if offset != end:
            return None
    entries["text/markdown"] = markdown
    payload = bytearray(struct.pack("<I", len(entries)))
    for key, value in entries.items():
        for text in (key, value):
            if len(text) > _MAX_BYTES // 2:
                return None
            data = text.encode("utf-16-le")
            if len(payload) + 8 + len(data) > _MAX_BYTES:
                return None
            payload.extend(struct.pack("<I", len(data) // 2))
            payload.extend(data)
            payload.extend(b"\0" * (-len(data) % 4))
    return struct.pack("<I", len(payload)) + payload
