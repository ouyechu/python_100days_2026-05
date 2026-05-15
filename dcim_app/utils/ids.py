from __future__ import annotations

import unicodedata


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip() == "" or v.strip().lower() == "nan"
    if isinstance(v, float) and v != v:  # NaN
        return True
    return False


def to_str(v) -> str | None:
    if is_blank(v):
        return None
    s = unicodedata.normalize("NFKC", str(v).strip())
    return s or None


def normalize_unique_id(val) -> str | None:
    """
    唯一ID（点分编号如 93.2014.2.7.1.1）：Excel/pandas 常读成 int、float、Decimal，
    若再 to_pickle 会失真；此处统一为 NFKC 后的字符串。
    """
    if is_blank(val):
        return None
    try:
        import numpy as np

        if isinstance(val, np.generic):
            val = val.item()
    except Exception:
        pass
    if isinstance(val, bool):
        return None
    if isinstance(val, int):
        return str(val)
    if isinstance(val, float):
        import math

        if math.isnan(val):
            return None
        if val.is_integer():
            return str(int(val))
        s = format(val, "f").rstrip("0").rstrip(".")
        return unicodedata.normalize("NFKC", s.strip()) or None
    try:
        from decimal import Decimal

        if isinstance(val, Decimal):
            if val != val:
                return None
            s = format(val, "f").rstrip("0").rstrip(".")
            return unicodedata.normalize("NFKC", s.strip()) or None
    except Exception:
        pass
    s = unicodedata.normalize("NFKC", str(val).strip())
    return s if s else None


UID_ROW_KEYS = ("唯一ID", "唯一Id", "unique_id", "UniqueID", "UNIQUE_ID")


def unique_id_from_row(row: dict) -> str | None:
    """从行中取唯一ID（兼容列名变体；列名需已 NFKC/trim）。"""
    if not isinstance(row, dict):
        return None
    for k in UID_ROW_KEYS:
        if k in row:
            uid = normalize_unique_id(row.get(k))
            if uid:
                return uid
    return None

