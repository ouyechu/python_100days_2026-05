from __future__ import annotations

from datetime import datetime

from flask_login import UserMixin

from .extensions import db, login_manager


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    display_name = db.Column(db.String(100))
    employee_id = db.Column(db.String(20))
    role = db.Column(db.String(20), default="user")
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    def set_password(self, password: str):
        from werkzeug.security import generate_password_hash

        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        from werkzeug.security import check_password_hash

        return check_password_hash(self.password_hash, password)


class DeviceCategory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), unique=True, nullable=False)
    display_name = db.Column(db.String(100))
    description = db.Column(db.Text)
    devices = db.relationship("Device", backref="category", lazy=True)


class Device(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    category_id = db.Column(db.Integer, db.ForeignKey("device_category.id"))

    scope_id = db.Column(db.String(20))
    unique_id = db.Column(db.String(50))
    instance_name = db.Column(db.String(200))
    asset_status = db.Column(db.String(50))
    location = db.Column(db.String(500))
    brand = db.Column(db.String(100))
    model = db.Column(db.String(100))

    warranty_years = db.Column(db.String(20))
    lifecycle = db.Column(db.String(50))
    commission_date = db.Column(db.Date)
    warranty_start = db.Column(db.Date)
    warranty_end = db.Column(db.Date)
    is_expired = db.Column(db.Boolean)

    energy_type = db.Column(db.String(50))
    threshold_range = db.Column(db.String(100))
    device_type = db.Column(db.String(100))
    rated_capacity = db.Column(db.String(50))

    qr_code = db.Column(db.String(200))
    it_rack_mount_status = db.Column(db.String(50))

    object_kind = db.Column(db.String(20))  # device | space | unknown
    profession = db.Column(db.String(20))  # 电气 | 暖通 | 弱电 | 消防 | 未知
    model_classification = db.Column(db.String(100))

    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    events = db.relationship(
        "AssetEvent", backref="device", lazy=True, cascade="all, delete-orphan"
    )


class AssetEvent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    device_id = db.Column(db.Integer, db.ForeignKey("device.id"), nullable=False, index=True)

    event_type = db.Column(db.String(50), nullable=False)
    title = db.Column(db.String(200))
    status = db.Column(db.String(50), default="done")
    occurred_at = db.Column(db.DateTime, default=datetime.now, index=True)

    created_by = db.Column(db.Integer, db.ForeignKey("user.id"))
    note = db.Column(db.Text)
    payload_json = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.now)


@login_manager.user_loader
def load_user(user_id: str):
    return User.query.get(int(user_id))

