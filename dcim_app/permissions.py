from __future__ import annotations

from flask_login import current_user


def is_admin() -> bool:
    return bool(
        current_user.is_authenticated and (current_user.role in ("admin", "区域总监"))
    )


def can_access_maintenance() -> bool:
    return is_admin()


def can_export() -> bool:
    if not current_user.is_authenticated:
        return False
    if is_admin():
        return True
    return current_user.role == "弱电专业"


def allowed_professions_for_current_user() -> set[str] | None:
    """
    返回允许访问的专业集合；None 表示全专业可见。
    """
    if not current_user.is_authenticated:
        return set()

    role = (current_user.role or "").strip()
    if role in (
        "admin",
        "区域总监",
        "园区经理",
        "设施主管",
        "弱电专业",
        "user",
        "",
    ):
        return None
    if role == "电气专业":
        return {"电气"}
    if role == "暖通专业":
        return {"暖通"}
    return None

