"""
Site branding settings - app name, login page logo, sidebar logo, favicon.

Backed by a single AppSetting row (id=1), created on first access with
defaults matching the logo.png this app shipped with. Exposed to every
template via the `app_settings` context processor in app/__init__.py, so
base.html and login.html can both use it without each route needing to
pass it explicitly.
"""
import os
import time

from flask import current_app, url_for
from werkzeug.utils import secure_filename

from app.extensions import db
from app.models import AppSetting

UPLOAD_SUBDIR = "uploads/branding"  # under the static folder, forward slashes for url_for
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "svg", "ico", "gif", "webp"}


def get_or_create_settings() -> AppSetting:
    settings = db.session.get(AppSetting, 1)
    if settings is None:
        settings = AppSetting(id=1, app_name="Design Calculations")
        db.session.add(settings)
        db.session.commit()
    return settings


def _asset_url(filename, default="images/logo.png"):
    if filename:
        return url_for("static", filename=f"{UPLOAD_SUBDIR}/{filename}")
    return url_for("static", filename=default)


# Shipped with the app - the FCC logo pulled from the original DP103207
# workbook, used for the report header until someone uploads a "Report
# logo" of their own in Settings.
DEFAULT_REPORT_LOGO = "images/design_calc/fcc_logo.jpeg"


def settings_context(settings: AppSetting) -> dict:
    """Plain dict of everything templates need - app_name plus resolved
    image URLs, falling back to the built-in logo.png (or, for the report
    logo, the default FCC logo) when nothing's been uploaded for a slot."""
    return {
        "app_name": settings.app_name,
        "login_logo_url": _asset_url(settings.login_logo),
        "sidebar_logo_url": _asset_url(settings.sidebar_logo),
        "favicon_url": _asset_url(settings.favicon),
        "report_logo_url": _asset_url(settings.report_logo, default=DEFAULT_REPORT_LOGO),
    }


def get_report_logo_path() -> str:
    """Absolute filesystem path to the logo PDF reports should use -
    the uploaded "Report logo" if set, otherwise the default FCC logo.
    (PDF generation needs a real file path, not a URL.)"""
    settings = get_or_create_settings()
    if settings.report_logo:
        path = os.path.join(current_app.static_folder, *UPLOAD_SUBDIR.split("/"), settings.report_logo)
        if os.path.exists(path):
            return path
    return os.path.join(current_app.static_folder, *DEFAULT_REPORT_LOGO.split("/"))


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_branding_file(file_storage, slot: str) -> str:
    """Saves an uploaded branding image under static/uploads/branding/,
    named "<slot>_<timestamp>.<ext>" so replacing an image gets a fresh
    filename/URL every time (no stale-browser-cache issues). Returns just
    the filename to store on the AppSetting row."""
    ext = file_storage.filename.rsplit(".", 1)[1].lower()
    filename = secure_filename(f"{slot}_{int(time.time())}.{ext}")

    upload_dir = os.path.join(current_app.static_folder, *UPLOAD_SUBDIR.split("/"))
    os.makedirs(upload_dir, exist_ok=True)
    file_storage.save(os.path.join(upload_dir, filename))
    return filename


def delete_branding_file(filename: str):
    if not filename:
        return
    path = os.path.join(current_app.static_folder, *UPLOAD_SUBDIR.split("/"), filename)
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
