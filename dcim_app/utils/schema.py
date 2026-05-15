from __future__ import annotations

from sqlalchemy import inspect, text

from ..extensions import db


def ensure_device_schema_sqlite():
    """SQLite 增量加列（兼容老库）。"""
    try:
        url = str(db.engine.url)
    except Exception:
        return
    if "sqlite" not in url:
        return

    insp = inspect(db.engine)
    cols = {c["name"] for c in insp.get_columns("device")}
    to_add = []
    if "it_rack_mount_status" not in cols:
        to_add.append(("it_rack_mount_status", "VARCHAR(50)"))
    if "object_kind" not in cols:
        to_add.append(("object_kind", "VARCHAR(20)"))
    if "profession" not in cols:
        to_add.append(("profession", "VARCHAR(20)"))
    if "model_classification" not in cols:
        to_add.append(("model_classification", "VARCHAR(100)"))
    if not to_add:
        return

    with db.engine.begin() as conn:
        for col, typ in to_add:
            conn.execute(text(f"ALTER TABLE device ADD COLUMN {col} {typ}"))

