from __future__ import annotations

import os

from ..extensions import db
from ..models import Device, DeviceCategory
from ..utils.io import read_pkl_rows
from ..utils.ids import to_str, unique_id_from_row


def _guess_category_from_filename(file_path: str) -> str:
    stem = os.path.splitext(os.path.basename(file_path))[0].strip()
    for sep in ("-", "－", "—", "_", " "):
        if sep in stem:
            parts = [p.strip() for p in stem.split(sep) if p and p.strip()]
            if parts:
                return parts[-1]
    return stem or "其他设备"


def _warehouse_stem_normalized(file_path: str) -> str:
    stem = os.path.splitext(os.path.basename(file_path))[0].strip()
    return stem.lower().replace(" ", "")


def is_r_it_cabinet_ac_pdu_warehouse_file(file_path: str) -> bool:
    n = _warehouse_stem_normalized(file_path)
    return "r楼" in n and "it机柜交流pdu" in n


def is_r_it_cabinet_only_warehouse_file(file_path: str) -> bool:
    n = _warehouse_stem_normalized(file_path)
    if "r楼" not in n or "it机柜" not in n:
        return False
    if "it机柜交流pdu" in n:
        return False
    return True


def canonical_category_name_for_warehouse_file(file_path: str) -> str:
    if is_r_it_cabinet_ac_pdu_warehouse_file(file_path):
        return "交流PDU"
    if is_r_it_cabinet_only_warehouse_file(file_path):
        return "IT机柜"
    return _guess_category_from_filename(file_path)


def iter_warehouse_pkl(warehouse_dir: str):
    if not warehouse_dir or not os.path.isdir(warehouse_dir):
        return []
    files = []
    for name in os.listdir(warehouse_dir):
        if name.lower().endswith(".pkl"):
            files.append(os.path.join(warehouse_dir, name))
    files.sort(key=lambda p: os.path.basename(p))
    return files


def recategorize_misplaced_rdapb_branch_pdu() -> int:
    cat_rpp = DeviceCategory.query.filter_by(name="交流列头柜").first()
    cat_pdu = DeviceCategory.query.filter_by(name="交流PDU").first()
    if not cat_rpp or not cat_pdu:
        return 0
    q = Device.query.filter(
        Device.category_id == cat_rpp.id,
        Device.instance_name.isnot(None),
        Device.instance_name.contains("RDAPB"),
    )
    n = 0
    for d in q.all():
        d.category_id = cat_pdu.id
        n += 1
    if n:
        db.session.commit()
    return n


def merge_it_cabinet_ac_pdu_category_alias() -> int:
    alias = DeviceCategory.query.filter_by(name="IT机柜交流PDU").first()
    main = DeviceCategory.query.filter_by(name="交流PDU").first()
    if not alias or not main or alias.id == main.id:
        return 0
    moved = 0
    for d in Device.query.filter_by(category_id=alias.id).all():
        d.category_id = main.id
        moved += 1
    db.session.delete(alias)
    db.session.commit()
    return moved


def apply_post_warehouse_category_fixes() -> dict:
    return {
        "rdapb_moved_to_ac_pdu": recategorize_misplaced_rdapb_branch_pdu(),
        "it_cabinet_pdu_alias_merged": merge_it_cabinet_ac_pdu_category_alias(),
    }


def _ensure_category(name: str) -> DeviceCategory:
    category = DeviceCategory.query.filter_by(name=name).first()
    if category:
        return category
    category = DeviceCategory(
        name=name,
        display_name=name,
        description=f"{name}设备台账",
    )
    db.session.add(category)
    db.session.commit()
    return category


def load_devices_from_warehouse(
    *,
    warehouse_dir: str,
    update_device_from_row,
    merge_it_cabinet_rows,
    normalize_it_rack_mount_status_for_it_cabinets,
    prune_missing: bool = False,
) -> dict:
    """
    从“仓库源目录”批量加载/更新设备数据（upsert）。
    说明与旧逻辑一致：按文件名/行内设备分类推断分类；按 (category_id, 唯一ID) upsert。
    """
    pkl_files = iter_warehouse_pkl(warehouse_dir)
    if not pkl_files:
        pf = apply_post_warehouse_category_fixes()
        nr = normalize_it_rack_mount_status_for_it_cabinets()
        return {
            "warehouse_dir": warehouse_dir,
            "files": 0,
            "inserted": 0,
            "updated": 0,
            "skipped": 0,
            "deleted": 0,
            "post_category_fixes": pf,
            "it_rack_mount_normalized": nr,
        }

    total_inserted = 0
    total_updated = 0
    total_skipped = 0
    total_deleted = 0

    for file_path in pkl_files:
        default_category_name = canonical_category_name_for_warehouse_file(file_path)
        is_r_it_pdu_pkl = is_r_it_cabinet_ac_pdu_warehouse_file(file_path)
        is_r_it_cabinet_pkl = is_r_it_cabinet_only_warehouse_file(file_path)
        try:
            rows = read_pkl_rows(file_path)
        except Exception:
            total_skipped += 1
            continue

        if is_r_it_cabinet_pkl:
            rows = merge_it_cabinet_rows(rows)

        file_seen_unique_ids_by_category: dict[int, set[str]] = {}

        for row in rows:
            row_category_name = to_str(row.get("设备分类")) or default_category_name or "其他设备"
            if is_r_it_pdu_pkl:
                row_category_name = "交流PDU"
            if is_r_it_cabinet_pkl:
                row_category_name = "IT机柜"

            category = _ensure_category(row_category_name)

            uid = unique_id_from_row(row)
            if uid:
                file_seen_unique_ids_by_category.setdefault(category.id, set()).add(uid)
                existing = Device.query.filter_by(category_id=category.id, unique_id=uid).first()
                if not existing and (is_r_it_pdu_pkl or is_r_it_cabinet_pkl):
                    existing = Device.query.filter_by(unique_id=uid).first()
                    if existing:
                        existing.category_id = category.id
                if existing:
                    update_device_from_row(existing, row)
                    total_updated += 1
                else:
                    device = Device(category_id=category.id)
                    update_device_from_row(device, row)
                    db.session.add(device)
                    total_inserted += 1
            else:
                device = Device(category_id=category.id)
                update_device_from_row(device, row)
                db.session.add(device)
                total_inserted += 1

        db.session.commit()

        if prune_missing:
            for cat_id, seen_ids in file_seen_unique_ids_by_category.items():
                q = Device.query.filter(
                    Device.category_id == cat_id,
                    Device.unique_id.isnot(None),
                    Device.unique_id != "",
                )
                to_delete = q.filter(~Device.unique_id.in_(list(seen_ids))).all()
                for d in to_delete:
                    db.session.delete(d)
                db.session.commit()
                total_deleted += len(to_delete)

    post_fixes = apply_post_warehouse_category_fixes()
    nr = normalize_it_rack_mount_status_for_it_cabinets()

    return {
        "warehouse_dir": warehouse_dir,
        "files": len(pkl_files),
        "inserted": total_inserted,
        "updated": total_updated,
        "skipped": total_skipped,
        "deleted": total_deleted,
        "post_category_fixes": post_fixes,
        "it_rack_mount_normalized": nr,
    }

