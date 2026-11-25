"""Encoding-safe reader for Access SaveAsText files (UTF-16LE / UTF-8 / CP932)."""

from __future__ import annotations

from pathlib import Path


def decode_access_bytes(data: bytes) -> str:
    """Decode bytes produced by Access SaveAsText into normalized LF text."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = data.decode("utf-16")
    elif data.startswith(b"\xef\xbb\xbf"):
        text = data.decode("utf-8-sig")
    else:
        for enc in ("utf-8", "cp932", "utf-16-le"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = data.decode("utf-8", errors="replace")

    return text.replace("\r\n", "\n").replace("\r", "\n")


def read_access_text_file(path: str | Path) -> str:
    """Read an Access SaveAsText file with automatic encoding detection."""
    return decode_access_bytes(Path(path).read_bytes())
