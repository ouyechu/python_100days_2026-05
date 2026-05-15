from __future__ import annotations

import unicodedata


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        s = v.strip()
        return s == "" or s.lower() == "nan"
    if isinstance(v, float) and v != v:  # NaN
        return True
    return False


def to_str(v) -> str | None:
    if is_blank(v):
        return None
    s = unicodedata.normalize("NFKC", str(v).strip())
    return s or None

