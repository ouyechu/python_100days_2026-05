from __future__ import annotations

from datetime import datetime


def parse_date(date_val):
    if date_val is None:
        return None
    if isinstance(date_val, datetime):
        return date_val.date()
    try:
        import datetime as _dt

        if isinstance(date_val, _dt.date) and not isinstance(date_val, _dt.datetime):
            return date_val
    except Exception:
        pass
    try:
        from dateutil import parser

        return parser.parse(str(date_val)).date()
    except Exception:
        return None


def parse_datetime(dt_val):
    if dt_val is None:
        return None
    if isinstance(dt_val, datetime):
        return dt_val
    try:
        from dateutil import parser

        return parser.parse(str(dt_val))
    except Exception:
        return None

