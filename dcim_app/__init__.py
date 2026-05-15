from __future__ import annotations

import os

from flask import Flask

from .extensions import db, login_manager


def create_app() -> Flask:
    app = Flask(__name__)

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    upload_folder = os.path.join(base_dir, "uploads")
    os.makedirs(upload_folder, exist_ok=True)

    app.config["BASE_DIR"] = base_dir
    app.config["UPLOAD_FOLDER"] = upload_folder
    app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
    app.config["SECRET_KEY"] = "DC-Asset-Manager-Secret-2024"
    app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///asset_manager.db"

    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = "login"

    from .utils.schema import ensure_device_schema_sqlite

    @app.before_request
    def _ensure_device_schema_once():
        if app.config.get("_device_schema_v1"):
            return
        try:
            ensure_device_schema_sqlite()
        except Exception:
            pass
        app.config["_device_schema_v1"] = True

    # register routes
    from .routes import (
        admin_ops,
        auth,
        dashboard,
        devices,
        export_tools,
        it_cabinet,
        maintenance,
        users,
    )

    auth.register(app)
    dashboard.register(app)
    it_cabinet.register(app)
    devices.register(app)
    maintenance.register(app)
    users.register(app)
    export_tools.register(app)
    admin_ops.register(app)

    return app


def init_db(app: Flask):
    from .services.bootstrap import init_db as _init

    _init(app)

