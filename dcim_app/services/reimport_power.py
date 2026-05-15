from __future__ import annotations

import unicodedata

from ..extensions import db
from ..models import Device, DeviceCategory
from ..utils.io import read_pkl_rows
from ..utils.ids import unique_id_from_row


def normalize_pkl_row_keys(row: dict) -> dict:
    """
    列名 NFKC + trim；并兼容中文表头乱码（mojibake）映射回标准字段名。
    """
    out = {}
    for k, v in row.items():
        if k is None:
            continue
        key = unicodedata.normalize("NFKC", str(k).strip())
        if key:
            out[key] = v

    legacy_key_map = {
        "��Χ���": "范围域号",
        "ΨһID": "唯一ID",
        "ʵ������": "实例名称",
        "�ʲ�״̬": "资产状态",
        "λ����Ϣ": "位置信息",
        "Ʒ��": "品牌",
        "�ͺ�": "型号",
        "�ϼ�״̬": "IT机柜上下架状态",
        "����״̬": "资产状态",
    }
    for lk, nk in legacy_key_map.items():
        if lk in out and nk not in out:
            out[nk] = out.get(lk)
    return out


def ensure_category(name: str, *, display_name: str | None = None) -> DeviceCategory:
    c = DeviceCategory.query.filter_by(name=name).first()
    if c:
        return c
    c = DeviceCategory(
        name=name,
        display_name=display_name or name,
        description=f"{name}设备台账",
    )
    db.session.add(c)
    db.session.commit()
    return c


def delete_devices_by_category_name(category_name: str) -> int:
    cat = DeviceCategory.query.filter_by(name=category_name).first()
    if not cat:
        return 0
    rows = Device.query.filter_by(category_id=cat.id).all()
    n = len(rows)
    for d in rows:
        db.session.delete(d)
    db.session.commit()
    return n


def load_power_pkl_rows(pkl_path: str) -> list[dict]:
    rows = read_pkl_rows(pkl_path)
    return [normalize_pkl_row_keys(r) for r in rows if isinstance(r, dict)]


def reimport_power_pkls(
    *,
    pdu_pkl_path: str,
    it_cabinet_pkl_path: str,
    update_device_from_row,
    pdu_uid_suffix: str = "1.1.7.4",
    it_uid_suffix: str = "2.7.1.1",
) -> dict:
    """
    清空并按“上电 pkl”重灌两类数据：
    - 交流PDU：唯一ID 必须以 pdu_uid_suffix 结尾
    - IT机柜：唯一ID 必须以 it_uid_suffix 结尾
    """
    cat_pdu = ensure_category("交流PDU")
    cat_it = ensure_category("IT机柜")

    deleted_pdu = delete_devices_by_category_name("交流PDU")
    deleted_it = delete_devices_by_category_name("IT机柜")

    pdu_rows = load_power_pkl_rows(pdu_pkl_path)
    it_rows = load_power_pkl_rows(it_cabinet_pkl_path)

    stats = {
        "deleted": {"交流PDU": deleted_pdu, "IT机柜": deleted_it},
        "imported": {"交流PDU": {"inserted": 0, "updated": 0}, "IT机柜": {"inserted": 0, "updated": 0}},
        "skipped": {
            "交流PDU": {"no_unique_id": 0, "suffix_mismatch": 0},
            "IT机柜": {"no_unique_id": 0, "suffix_mismatch": 0},
        },
        "source": {"交流PDU": pdu_pkl_path, "IT机柜": it_cabinet_pkl_path},
        "suffix": {"交流PDU": pdu_uid_suffix, "IT机柜": it_uid_suffix},
        "rows": {"交流PDU": len(pdu_rows), "IT机柜": len(it_rows)},
    }

    def _import_rows(rows: list[dict], *, category: DeviceCategory, suffix: str, key: str):
        for row in rows:
            uid = unique_id_from_row(row)
            if not uid:
                stats["skipped"][key]["no_unique_id"] += 1
                continue
            if not str(uid).endswith(suffix):
                stats["skipped"][key]["suffix_mismatch"] += 1
                continue

            existing = Device.query.filter_by(category_id=category.id, unique_id=uid).first()
            if existing:
                update_device_from_row(existing, row)
                stats["imported"][key]["updated"] += 1
            else:
                d = Device(category_id=category.id)
                update_device_from_row(d, row)
                d.unique_id = uid
                db.session.add(d)
                stats["imported"][key]["inserted"] += 1

        db.session.commit()

    _import_rows(pdu_rows, category=cat_pdu, suffix=pdu_uid_suffix, key="交流PDU")
    _import_rows(it_rows, category=cat_it, suffix=it_uid_suffix, key="IT机柜")

    return {"ok": True, "stats": stats}

