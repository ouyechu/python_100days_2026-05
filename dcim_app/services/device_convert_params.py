from __future__ import annotations

import os
import unicodedata

from ..utils.io import read_pkl_rows
from ..utils.ids import normalize_unique_id, to_str, unique_id_from_row

# 列名兼容（历史 GBK 误读等）
_PARAM_KEYS = (
    "参数",
    "参 数",
)


def _normalize_model(val) -> str | None:
    s = to_str(val)
    return s


def _param_from_row(row: dict) -> str | None:
    if not isinstance(row, dict):
        return None
    for k in _PARAM_KEYS:
        if k in row:
            s = to_str(row.get(k))
            if s:
                return s
    for key in row:
        if not isinstance(key, str):
            continue
        nk = unicodedata.normalize("NFKC", key.strip())
        if nk == "参数" or (nk.endswith("参数") and len(nk) <= 8):
            s = to_str(row.get(key))
            if s:
                return s
    return None


def _model_from_row(row: dict) -> str | None:
    if not isinstance(row, dict):
        return None
    for k in ("型号", "型 号"):
        if k in row:
            m = _normalize_model(row.get(k))
            if m:
                return m
    return None


def _scan_mtime(warehouse_dir: str) -> float:
    if not os.path.isdir(warehouse_dir):
        return -1.0
    mt = 0.0
    for name in os.listdir(warehouse_dir):
        if not name.lower().endswith(".pkl"):
            continue
        p = os.path.join(warehouse_dir, name)
        try:
            mt = max(mt, os.path.getmtime(p))
        except OSError:
            continue
    return mt


_cache_mtime: float | None = None
_cache_index: dict[tuple[str, str], tuple[str, str]] | None = None


def _build_index(warehouse_dir: str) -> dict[tuple[str, str], tuple[str, str]]:
    """
    (唯一ID, 型号) -> (参数字符串, pkl 文件名)
    后读到的行覆盖先读到的（与同名键最后一次写入一致）。
    """
    idx: dict[tuple[str, str], tuple[str, str]] = {}
    if not os.path.isdir(warehouse_dir):
        return idx
    names = sorted(
        n for n in os.listdir(warehouse_dir) if n.lower().endswith(".pkl")
    )
    for name in names:
        path = os.path.join(warehouse_dir, name)
        try:
            rows = read_pkl_rows(path)
        except Exception:
            continue
        for row in rows:
            uid = unique_id_from_row(row)
            mod = _model_from_row(row)
            if not uid or not mod:
                continue
            param = _param_from_row(row)
            if not param:
                continue
            idx[(uid, mod)] = (param, name)
    return idx


def get_convert_param_index(*, warehouse_dir: str) -> dict[tuple[str, str], tuple[str, str]]:
    global _cache_mtime, _cache_index
    mt = _scan_mtime(warehouse_dir)
    if _cache_index is not None and _cache_mtime == mt:
        return _cache_index
    _cache_index = _build_index(warehouse_dir)
    _cache_mtime = mt
    return _cache_index


def lookup_convert_params(
    *,
    warehouse_dir: str,
    unique_id: str | None,
    model: str | None,
) -> dict:
    uid = normalize_unique_id(unique_id) if unique_id else None
    mod = _normalize_model(model)
    out: dict = {
        "matched": False,
        "parameter_text": None,
        "source_file": None,
        "unique_id": uid,
        "model": mod,
    }
    if not uid or not mod:
        return out
    idx = get_convert_param_index(warehouse_dir=warehouse_dir)
    hit = idx.get((uid, mod))
    if not hit:
        return out
    text, src = hit
    out["matched"] = True
    out["parameter_text"] = text
    out["source_file"] = src
    return out
