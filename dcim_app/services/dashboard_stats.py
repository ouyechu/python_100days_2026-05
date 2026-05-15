from __future__ import annotations

import re
from collections import Counter, defaultdict

from sqlalchemy import or_

from ..enums import ASSET_STATUS_ENUM
from ..models import Device, DeviceCategory
from ..utils.ids import is_blank


PDU_PHASE_SUFFIX_RE = re.compile(r"-[ABC]\([^)]+\)\s*$")
PDU_CABINET_CODE_RE = re.compile(r"\(([^)]+)\)\s*$")


def base_pdu_branch_name(instance_name: str | None):
    if is_blank(instance_name):
        return None
    s = str(instance_name).strip()
    m = PDU_PHASE_SUFFIX_RE.search(s)
    return s[: m.start()] if m else s


def parse_pdu_cabinet_code(instance_name: str | None):
    if is_blank(instance_name):
        return None
    m = PDU_CABINET_CODE_RE.search(str(instance_name).strip())
    return m.group(1).strip() if m else None


def is_r_it_cabinet_ac_pdu_category(name: str | None) -> bool:
    if not name:
        return False
    n = str(name).strip()
    if n == "交流PDU":
        return True
    if "IT机柜" in n and "交流PDU" in n:
        return True
    return False


def is_it_cabinet_ledger_category(name: str | None) -> bool:
    if not name:
        return False
    n = str(name).strip()
    if "IT机柜" not in n:
        return False
    return not is_r_it_cabinet_ac_pdu_category(n)


def normalize_asset_status_bucket(status: str | None) -> str:
    if is_blank(status):
        return "未标注"
    s = str(status).strip()
    if s.startswith("使用中"):
        return "使用中"
    if s.startswith("停用"):
        return "停用中"
    if s.startswith("故障"):
        return "故障中"
    if s.startswith("维修"):
        return "维修中"
    if "报废" in s:
        return "已报废"
    if s in ASSET_STATUS_ENUM:
        return s
    return s


def rollup_cabinet_status(norm_set: set) -> str:
    if not norm_set:
        return "未标注"
    if "使用中" in norm_set and "停用中" in norm_set:
        return "混合"
    if "使用中" in norm_set:
        return "使用中"
    if "停用中" in norm_set:
        return "停用中"
    if len(norm_set) == 1:
        return next(iter(norm_set))
    return "混合"


def compute_it_cabinet_dashboard():
    """
    IT 机柜使用状态 + R楼-IT机柜交流 PDU 数量（按三相合并为一条支路重算）。
    """
    pdu_category_ids = [
        c.id for c in DeviceCategory.query.all() if is_r_it_cabinet_ac_pdu_category(c.name)
    ]
    pdu_category_names = [
        c.name for c in DeviceCategory.query.all() if is_r_it_cabinet_ac_pdu_category(c.name)
    ]

    ledger_cat_ids = [
        c.id for c in DeviceCategory.query.all() if is_it_cabinet_ledger_category(c.name)
    ]

    if ledger_cat_ids:
        q_led = Device.query.filter(Device.category_id.in_(ledger_cat_ids))
        led_devices = q_led.all()
        led_counts = Counter(normalize_asset_status_bucket(d.asset_status) for d in led_devices)
        cabinet_total = len(led_devices)
        cabinet_by_status = [
            {"label": k, "count": v}
            for k, v in sorted(led_counts.items(), key=lambda x: (-x[1], x[0]))
        ]
        source = "ledger"
    else:
        source = "pdu_derived"
        cabinet_total = 0
        cabinet_by_status = []

    phase_rows = []
    if pdu_category_ids:
        q_pdu = Device.query.filter(
            Device.category_id.in_(pdu_category_ids),
            Device.location.isnot(None),
        ).filter(
            or_(
                Device.location.contains("楼R"),
                Device.location.contains("/R/"),
                Device.location.contains("\\R\\"),
            )
        )
        q_pdu = q_pdu.filter(
            or_(
                Device.location.contains("IT机房"),
                Device.location.contains("_IT机房"),
            )
        )
        phase_rows = q_pdu.all()

    branch_keys = set()
    cab_norms: dict[str, set] = defaultdict(set)
    for d in phase_rows:
        b = base_pdu_branch_name(d.instance_name)
        if b:
            branch_keys.add(b)
        code = parse_pdu_cabinet_code(d.instance_name)
        if code:
            cab_norms[code].add(normalize_asset_status_bucket(d.asset_status))

    branch_count = len(branch_keys)
    phase_count = len(phase_rows)

    if source == "pdu_derived":
        cabinet_total = len(cab_norms)
        roll_counts = Counter(rollup_cabinet_status(s) for s in cab_norms.values())
        cabinet_by_status = [
            {"label": k, "count": v}
            for k, v in sorted(roll_counts.items(), key=lambda x: (-x[1], x[0]))
        ]

    return {
        "source": source,
        "cabinet_total": cabinet_total,
        "cabinet_by_status": cabinet_by_status,
        "pdu": {
            "category_names": pdu_category_names,
            "phase_count": phase_count,
            "branch_count": branch_count,
            "cabinet_codes_from_pdu": len(cab_norms),
        },
    }

