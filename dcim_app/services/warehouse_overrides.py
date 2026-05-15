from __future__ import annotations

import os

from ..utils.ids import to_str, unique_id_from_row


def _pkl_asset_status_in_use(st: str | None) -> bool:
    s = to_str(st)
    if not s:
        return False
    return str(s).strip().startswith("使用中")


def _dedupe_pkl_rows_by_unique_id(rows: list) -> list[dict]:
    by_uid: dict[str, dict] = {}
    no_uid: list[dict] = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        uid = unique_id_from_row(r)
        if uid:
            by_uid[uid] = r
        else:
            no_uid.append(r)
    return list(by_uid.values()) + no_uid


def load_ac_pdu_override_from_pkl(*, warehouse_dir: str, read_pkl_rows) -> dict | None:
    """
    交流PDU数量以仓库源 `R楼-交流PDU.pkl` 为准（若存在）。
    返回：{'total': int, 'active': int, 'path': str}
    """
    if not warehouse_dir:
        return None
    p = os.path.join(warehouse_dir, "R楼-交流PDU.pkl")
    if not os.path.exists(p):
        return None
    try:
        rows = read_pkl_rows(p)
    except Exception:
        return None
    rows = _dedupe_pkl_rows_by_unique_id(rows)
    total = len(rows)
    active = sum(1 for row in rows if _pkl_asset_status_in_use(row.get("资产状态")))
    return {"total": total, "active": active, "path": p}


def load_it_cabinet_override_from_pkl(*, warehouse_dir: str, read_pkl_rows, merge_it_cabinet_rows) -> dict | None:
    """
    首页 IT 机柜数量以 `R楼-IT机柜.pkl` 为准（若存在且已放入仓库目录）。
    """
    if not warehouse_dir:
        return None
    p = os.path.join(warehouse_dir, "R楼-IT机柜.pkl")
    if not os.path.exists(p):
        return None
    try:
        rows = read_pkl_rows(p)
    except Exception:
        return None
    rows = merge_it_cabinet_rows(rows)
    total = len(rows)
    active = 0
    for row in rows:
        if not isinstance(row, dict):
            continue
        if _pkl_asset_status_in_use(row.get("资产状态")):
            active += 1
            continue
        # 若台账另有上下架列，只有已上架算使用
        st = to_str(row.get("IT机柜上下架状态") or row.get("上下架状态") or row.get("上架状态"))
        if st == "已上架":
            active += 1
    return {"total": total, "active": active, "path": p}

