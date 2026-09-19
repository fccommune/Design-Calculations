import io
import json
import os
import threading
import zipfile
from functools import wraps

from flask import render_template, request, jsonify, send_file, abort, current_app, flash, redirect, url_for
from flask_login import login_required, current_user
from sqlalchemy import func, or_
from werkzeug.security import generate_password_hash

from app.blueprints.dashboard import dashboard_bp
from app.extensions import db
from app.models import User
from app.services.app_settings import (
    get_or_create_settings,
    allowed_file as allowed_branding_file,
    save_branding_file,
    delete_branding_file,
)
from app.services.excel_import import parse_excel_rows, ExcelImportError
from app.services.gad_field_options import get_field_options, FIELD_SOURCES
from app.services.gad_globe_lookup import cell, resolve_globe_gad, GadLookupError
from app.services.gad_masters_import import (
    parse_single_table,
    MasterImportError,
    MODELS_BY_TABLE,
    TABLE_LABELS,
    EXCEL_HEADERS_BY_TABLE,
)
from app.services.solidworks_automation import (
    generate_globe_gad,
    solidworks_diagnostics,
    GadGenerationError,
    SolidWorksNotInstalledError,
    GENERATION_TOTAL_STEPS,
)
from app.services import gad_progress
from app.services.design_calc_report import build_24in_600_report


def admin_required(view):
    """Blocks non-admin users from user-management routes - creating
    accounts (and, implicitly, granting the admin role itself) is
    sensitive enough that it shouldn't be open to every logged-in user."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if current_user.role != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped


@dashboard_bp.route("/dashboard")
@login_required
def dashboard():
    user_count = User.query.count()
    active_count = User.query.filter_by(status=True).count()
    return render_template(
        "dashboard.html",
        user_count=user_count,
        active_count=active_count,
    )


@dashboard_bp.route("/users")
@login_required
@admin_required
def users():
    all_users = User.query.order_by(User.created_at.desc()).all()
    return render_template("users.html", users=all_users)


@dashboard_bp.route("/users/create", methods=["POST"])
@login_required
@admin_required
def users_create():
    payload = request.get_json(silent=True) or {}
    username = (payload.get("username") or "").strip()
    email = (payload.get("email") or "").strip()
    displayname = (payload.get("displayname") or "").strip()
    password = payload.get("password") or ""
    role = payload.get("role") or "user"
    active = bool(payload.get("active", True))

    if not username or not email or not password:
        return jsonify({"error": "Username, email and password are required."}), 400
    if len(password) < 6:
        return jsonify({"error": "Password must be at least 6 characters."}), 400
    if role not in ("user", "admin"):
        return jsonify({"error": "Invalid role."}), 400

    existing = User.query.filter(
        or_(func.lower(User.username) == username.lower(), func.lower(User.email) == email.lower())
    ).first()
    if existing:
        field = "username" if existing.username.lower() == username.lower() else "email"
        return jsonify({"error": f"That {field} is already in use."}), 409

    user = User(
        username=username,
        email=email,
        displayname=displayname or username,
        password=generate_password_hash(password, method="pbkdf2:sha256", salt_length=8),
        role=role,
        status=active,
    )
    db.session.add(user)
    db.session.commit()

    return jsonify({
        "id": user.id,
        "username": user.username,
        "email": user.email,
        "displayname": user.displayname,
        "role": user.role,
        "status": user.status,
        "created_at": user.created_at.strftime("%d-%b-%Y"),
    }), 201


@dashboard_bp.route("/settings", methods=["GET", "POST"])
@login_required
@admin_required
def settings():
    """Site branding - app name, login page logo, sidebar logo, favicon.
    See app.services.app_settings for how these get resolved into URLs for
    every template via the app_settings context processor."""
    setting = get_or_create_settings()

    if request.method == "POST":
        app_name = request.form.get("app_name", "").strip()
        if app_name:
            setting.app_name = app_name

        for slot, label in (
            ("login_logo", "Login page logo"),
            ("sidebar_logo", "Sidebar logo"),
            ("favicon", "Favicon"),
            ("report_logo", "Report logo"),
        ):
            uploaded = request.files.get(slot)
            if uploaded and uploaded.filename:
                if not allowed_branding_file(uploaded.filename):
                    flash(f"{label}: unsupported file type.", "danger")
                    continue
                old_filename = getattr(setting, slot)
                setattr(setting, slot, save_branding_file(uploaded, slot))
                delete_branding_file(old_filename)

        db.session.commit()
        flash("Settings saved.", "success")
        return redirect(url_for("dashboard.settings"))

    return render_template("settings.html", setting=setting)


@dashboard_bp.route("/globe")
@login_required
def globe():
    return render_template("globe.html")


@dashboard_bp.route("/globe/upload", methods=["POST"])
@login_required
def globe_upload():
    """Parses the uploaded valve BOM workbook and returns it as JSON so the
    page can render it as a table - equivalent to the old app's
    LoadDataFromExcel binding a DataTable to the DataGridView."""
    uploaded = request.files.get("excel_file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Please choose an Excel file."}), 400

    if not uploaded.filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
        return jsonify({"error": "Please upload a .xls, .xlsx or .xlsm file."}), 400

    try:
        headers, rows = parse_excel_rows(uploaded.stream)
    except ExcelImportError as exc:
        return jsonify({"error": str(exc)}), 400

    if not rows:
        return jsonify({"error": "No data rows found in the first sheet."}), 400

    return jsonify({"headers": headers, "rows": rows})


@dashboard_bp.route("/globe/field-options", methods=["POST"])
@login_required
def globe_field_options():
    """Valid master-table values for each of the given BOM column names, so
    the row editor can render those fields as dropdowns and flag a value
    that doesn't match any known master row. Fields with no master-table
    source are simply absent from the response (row editor keeps those as
    plain text)."""
    payload = request.get_json(silent=True) or {}
    fields = payload.get("fields")
    if not isinstance(fields, list):
        return jsonify({"error": "Expected a JSON list of field names."}), 400

    options = {}
    for field in fields:
        if not isinstance(field, str) or field not in FIELD_SOURCES:
            continue
        values = get_field_options(field)
        if values:
            options[field] = values

    return jsonify({"options": options})


@dashboard_bp.route("/globe/solidworks-check")
@login_required
def globe_solidworks_check():
    """Diagnostic page - reports exactly what the SolidWorks-installed check
    sees on whichever machine actually runs this Flask process (registry
    results per view, hostname, Python bitness). Visit this in a browser
    from the machine in question rather than needing shell/remote access to
    it - useful when 'SolidWorks Not Found' shows up unexpectedly."""
    return jsonify(solidworks_diagnostics())


@dashboard_bp.route("/globe/generate", methods=["POST"])
@login_required
def globe_generate():
    """Resolves the master-table lookups for one BOM row (drawing/hookup/
    cross-section numbers, dimension values, all title-block properties -
    see resolve_globe_gad) right here on the server, then packages the
    result into a small downloadable zip: the resolved job.json plus a
    pre-built, self-contained GadGenerate.exe. The browser downloads that
    zip; the user unzips it on their own machine (which needs SolidWorks
    installed, but not this web app, Python, or any database access) and
    double-clicks the exe to actually drive SolidWorks and produce the PDF.

    This is why resolution happens here rather than in the exe: it needs
    the database, and doing it now means a bad row (no matching master
    data) is reported immediately, instead of only after downloading and
    running something."""
    payload = request.get_json(silent=True) or {}
    row = payload.get("row")
    if not isinstance(row, dict):
        return jsonify({"error": "No row selected."}), 400

    series = cell(row, "Series")
    if series not in ("10", "11", "20", "21"):
        return jsonify({"error": "Please fill a correct Valve Series."}), 400
    if series in ("20", "21"):
        return jsonify({"error": "Butterfly valve (Series 20/21) generation isn't implemented yet - only Globe (10/11)."}), 400

    try:
        resolved = resolve_globe_gad(row)
    except GadLookupError as exc:
        return jsonify({"error": str(exc)}), 422

    downloads_dir = os.path.join(current_app.static_folder, "downloads")
    exe_path = os.path.join(downloads_dir, "GadGenerate.exe")
    if not os.path.exists(exe_path):
        return jsonify({"error": "GadGenerate.exe isn't available on the server yet - contact your admin."}), 500

    dwg_no = (resolved.get("properties") or {}).get("DWG_NO") or "GAD"
    safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in dwg_no)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("job.json", json.dumps(resolved, indent=2))
        zf.write(exe_path, "GadGenerate.exe")
        env_example_path = os.path.join(downloads_dir, ".env.example")
        if os.path.exists(env_example_path):
            zf.write(env_example_path, ".env.example")
        readme_path = os.path.join(downloads_dir, "README.txt")
        if os.path.exists(readme_path):
            zf.write(readme_path, "README.txt")
    buffer.seek(0)

    return send_file(
        buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"{safe_name}_GAD_package.zip",
    )


@dashboard_bp.route("/globe/generate-server", methods=["POST"])
@login_required
def globe_generate_server():
    """Kicks off PDF generation directly on whichever machine runs this
    Flask process - the "single dedicated server" option: only works if
    that machine itself has SolidWorks installed and licensed (e.g. the
    whole app deployed on .167). Alternative to /globe/generate's download-
    and-run-locally package for that deployment shape, where there's no
    need to download anything since everyone's browser already points at
    the SolidWorks machine.

    Runs the actual generation in a background thread rather than blocking
    this request, so the browser can poll /globe/generate-server/<token>/
    progress for real step-by-step status (see globe.html) instead of
    sitting on one long spinner. This also means concurrent clicks no
    longer serialize "for free" via one blocked Flask request at a time -
    generate_from_resolved() now takes an explicit lock for that instead
    (see its docstring), since SolidWorks COM automation is still only
    one-at-a-time regardless of how many browser tabs are waiting."""
    payload = request.get_json(silent=True) or {}
    row = payload.get("row")
    if not isinstance(row, dict):
        return jsonify({"error": "No row selected."}), 400

    series = cell(row, "Series")
    if series not in ("10", "11", "20", "21"):
        return jsonify({"error": "Please fill a correct Valve Series."}), 400
    if series in ("20", "21"):
        return jsonify({"error": "Butterfly valve (Series 20/21) generation isn't implemented yet - only Globe (10/11)."}), 400

    dwg_no = cell(row, "Dwg No") or "GA_Drawing"
    download_name = f"{dwg_no}.pdf".replace("/", "-")
    token = gad_progress.start_job(GENERATION_TOTAL_STEPS)
    close_event = gad_progress.get_close_event(token)

    # The background thread has none of this request's Flask context - it
    # needs the real app object (not the request-bound proxy) to push its
    # own app context, the same pattern worker.py uses for its poll loop.
    app_obj = current_app._get_current_object()

    def run_job():
        with app_obj.app_context():
            try:
                def on_progress(step_text, step_index):
                    gad_progress.update(token, step_text, step_index)

                def on_pdf_ready(pdf_path):
                    # Marks the job downloadable immediately - the thread
                    # keeps running after this to hold SolidWorks open until
                    # close_event is set (see /generate-server/<token>/close)
                    # or its own safety timeout elapses.
                    gad_progress.finish(token, pdf_path, download_name)

                generate_globe_gad(
                    row,
                    progress_cb=on_progress,
                    on_pdf_ready=on_pdf_ready,
                    close_event=close_event,
                )
            except SolidWorksNotInstalledError as exc:
                gad_progress.fail(token, str(exc))
            except (GadLookupError, GadGenerationError) as exc:
                gad_progress.fail(token, str(exc))
            except Exception as exc:  # unexpected SolidWorks/COM failure
                app_obj.logger.exception("GAD generation failed")
                gad_progress.fail(token, f"GAD generation failed: {exc}")

    threading.Thread(target=run_job, daemon=True).start()
    return jsonify({"token": token}), 202


@dashboard_bp.route("/globe/generate-server/<token>/progress")
@login_required
def globe_generate_server_progress(token):
    job = gad_progress.get_job(token)
    if job is None:
        return jsonify({"error": "Unknown or expired job."}), 404
    return jsonify({
        "step": job["step"],
        "step_index": job["step_index"],
        "total_steps": job["total_steps"],
        "done": job["done"],
        "error": job["error"],
    })


@dashboard_bp.route("/globe/generate-server/<token>/download")
@login_required
def globe_generate_server_download(token):
    job = gad_progress.get_job(token)
    if job is None or not job["done"] or job["error"] or not job["pdf_path"]:
        abort(404)
    return send_file(
        job["pdf_path"],
        mimetype="application/pdf",
        as_attachment=True,
        download_name=job["download_name"],
    )


@dashboard_bp.route("/globe/generate-server/<token>/close", methods=["POST"])
@login_required
def globe_generate_server_close(token):
    """Tells that job's still-running generation thread to close SolidWorks
    now (discarding any unsaved changes, no prompt) - called once the
    browser's "PDF generated successfully" message is dismissed, so
    SolidWorks doesn't sit open on the server forever after every
    generation. See generate_from_resolved()'s close_event docstring for
    why this has to be a signal to that same thread rather than closing it
    directly here."""
    gad_progress.request_close(token)
    return jsonify({"ok": True})


@dashboard_bp.route("/globe-calculation/body-bonnet-bolting")
@login_required
def body_bonnet_bolting():
    return render_template("body_bonnet_bolting.html")


@dashboard_bp.route("/globe-calculation/body-bonnet-bolting-draft")
@login_required
def body_bonnet_bolting_draft():
    return render_template("body_bonnet_bolting_draft.html")


@dashboard_bp.route("/globe-calculation/design-calc-24in-600")
@login_required
def design_calc_24in_600():
    return render_template("design_calc_24in_600.html")


@dashboard_bp.route("/globe-calculation/design-calc-24in-600/report", methods=["POST"])
@login_required
def design_calc_24in_600_report():
    """Formats whatever the browser already computed (sent up as a flat
    {field_id: value} dict) into a downloadable PDF - see
    app.services.design_calc_report for why nothing is recalculated here."""
    payload = request.get_json(silent=True) or {}
    if not isinstance(payload, dict):
        return jsonify({"error": "Expected a JSON object of field values."}), 400

    buffer = build_24in_600_report(payload)
    return send_file(
        buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="24in 600# Design Calculation Report.pdf",
    )


@dashboard_bp.route("/gad-masters")
@login_required
def gad_masters():
    # Fixed 2-column display order (fills column 1 top-to-bottom, then
    # column 2) rather than MODELS_BY_TABLE's natural order - matches how
    # these 5 tables are grouped in the old WinForms app's tooling.
    display_order = [
        "ga_globe_table", "ga_dim_valve_globe", "ga_globe_hookup",
        "ga_crosssec_globe", "ga_dim_act_globe",
    ]
    tables = [
        {"name": name, "label": TABLE_LABELS[name], "count": MODELS_BY_TABLE[name].query.count()}
        for name in display_order
    ]
    return render_template("gad_masters.html", tables=tables)


@dashboard_bp.route("/gad-masters/<table_name>")
@login_required
def gad_masters_table(table_name):
    if table_name not in MODELS_BY_TABLE:
        abort(404)
    model = MODELS_BY_TABLE[table_name]
    return render_template(
        "gad_masters_table.html",
        table_name=table_name,
        label=TABLE_LABELS[table_name],
        count=model.query.count(),
        expected_headers=EXCEL_HEADERS_BY_TABLE[table_name],
    )


@dashboard_bp.route("/gad-masters/<table_name>/upload", methods=["POST"])
@login_required
def gad_masters_table_upload(table_name):
    """Bulk-loads one master lookup table from its own uploaded workbook."""
    if table_name not in MODELS_BY_TABLE:
        abort(404)
    model = MODELS_BY_TABLE[table_name]

    uploaded = request.files.get("excel_file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Please choose an Excel file."}), 400

    if not uploaded.filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
        return jsonify({"error": "Please upload a .xls, .xlsx or .xlsm file."}), 400

    replace = request.form.get("replace") == "1"

    try:
        rows = parse_single_table(uploaded.stream, table_name)
    except MasterImportError as exc:
        return jsonify({"error": str(exc)}), 400

    if not rows:
        return jsonify({"error": "No data rows found in the sheet."}), 400

    if replace:
        model.query.delete()
    db.session.bulk_insert_mappings(model, rows)
    db.session.commit()

    return jsonify({"inserted": len(rows), "total": model.query.count()})


def _parse_filters_param(raw, columns, exclude=None):
    """Parses the `filters` query param - a JSON object of
    {column: [selected values]} - dropping anything that isn't a known
    column (or is `exclude`, when computing another column's own distinct
    list) or isn't a non-empty list."""
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
    except (ValueError, TypeError):
        return {}
    if not isinstance(parsed, dict):
        return {}
    return {
        col: values for col, values in parsed.items()
        if col in columns and col != exclude and isinstance(values, list) and values
    }


def _apply_filters(query, model, filters):
    for col, values in filters.items():
        query = query.filter(getattr(model, col).in_(values))
    return query


@dashboard_bp.route("/gad-masters/<table_name>/rows")
@login_required
def gad_masters_table_rows(table_name):
    """Paginated JSON dump of one master table's rows, for the on-page data
    viewer (tables can run into the thousands of rows, so this is server-
    side paged rather than sent all at once). Accepts an optional `filters`
    query param - a JSON object of {column: [selected values]} - for the
    Excel-style column filters."""
    if table_name not in MODELS_BY_TABLE:
        abort(404)
    model = MODELS_BY_TABLE[table_name]

    try:
        page = max(int(request.args.get("page", 1)), 1)
    except ValueError:
        page = 1
    try:
        per_page = min(max(int(request.args.get("per_page", 50)), 1), 200)
    except ValueError:
        per_page = 50

    columns = [c.name for c in model.__table__.columns if c.name != "id"]
    filters = _parse_filters_param(request.args.get("filters"), columns)

    base_query = _apply_filters(model.query, model, filters)
    total = base_query.count()
    records = (
        base_query.order_by(model.id)
        .offset((page - 1) * per_page)
        .limit(per_page)
        .all()
    )
    rows = [
        {c: ("" if getattr(r, c) is None else str(getattr(r, c))) for c in columns}
        for r in records
    ]

    return jsonify({
        "columns": columns,
        "rows": rows,
        "page": page,
        "per_page": per_page,
        "total": total,
        "pages": max((total + per_page - 1) // per_page, 1),
    })


@dashboard_bp.route("/gad-masters/<table_name>/distinct")
@login_required
def gad_masters_table_distinct(table_name):
    """Distinct values for one column, for that column's filter dropdown -
    scoped by any other active column filters (so choices narrow the same
    way Excel's AutoFilter dropdowns do)."""
    if table_name not in MODELS_BY_TABLE:
        abort(404)
    model = MODELS_BY_TABLE[table_name]
    columns = [c.name for c in model.__table__.columns if c.name != "id"]

    column = request.args.get("column")
    if column not in columns:
        return jsonify({"error": "Unknown column."}), 400

    filters = _parse_filters_param(request.args.get("filters"), columns, exclude=column)
    col_attr = getattr(model, column)
    query = _apply_filters(db.session.query(col_attr).distinct(), model, filters)
    values = ["" if v[0] is None else str(v[0]) for v in query.order_by(col_attr).all()]

    return jsonify({"values": values})
