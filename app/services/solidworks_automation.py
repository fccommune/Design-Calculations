"""
Drives SolidWorks over COM to assemble a Globe valve (Series 10/11) GAD
drawing and export it to PDF - a Python/pywin32 port of GenerateSheet1/2/3
in the old WinForms app's Form2.cs.

Same overall approach as the original: open one shared .DRWDOT template,
insert pre-built .SLDBLK blocks (GA block, hookup schematic) onto Sheet1,
push ~50 custom properties for the title block, insert the nameplate block
onto Sheet3, rebuild, export all 3 sheets to one PDF. Sheet2 is different:
it no longer holds a .SLDBLK cross-section block - it's replaced wholesale
with a whole separate .SLDDRW reference drawing, chosen per-row from the
ga_sheet4_globe master table (see
app.services.gad_globe_lookup.find_sheet4_drawing_no) - see
_generate_sheet2()'s docstring for why that copy needs a different mechanism
(this table/lookup is still named after "Sheet4" - that's what this content
used to be positioned as before it replaced the old Sheet2 outright; nothing
about the master table or its GAD Masters page needed renaming for this).

Runs synchronously in the Flask request (this app is expected to run on the
same Windows machine that has SolidWorks installed and licensed - COM
automation of a GUI app like SolidWorks can't run on a remote/Linux host).

Requires: pip install pywin32, and SolidWorks installed + its type library
registered - see _connect()/_ensure_typelibs() below for how the generated
wrapper (which is what exposes the swconst enum names as
win32com.client.constants.* below) actually gets built.
"""
import os
import tempfile
import threading
import winreg

import pythoncom
import win32com
import win32com.client

# win32com.client.gencache generates/caches a Python wrapper for a COM type
# library on first use, normally under the pywin32 install's own
# site-packages folder. If this Python was installed without admin rights,
# site-packages is read-only for this user and that write fails. Redirect
# the cache to a location that's always writable, same fix worker.py already
# applies for its frozen-exe case.
win32com.__gen_path__ = os.path.join(tempfile.gettempdir(), "gen_py")

from win32com.client import constants as sw_const

from app.services.gad_globe_lookup import cell
from app.services import gad_globe_lookup as lookup

BLOCK_INSERT_SCALE = 0.1

# Fixed pipeline of user-visible milestones, in order - shared with the
# frontend's progress UI (see globe.html) so the browser can show real
# step-by-step status instead of one long spinner while this runs. Index
# into this list is what gets reported to progress_cb(step_text, index)
# below; GENERATION_TOTAL_STEPS is exposed for the route that starts the
# progress-tracking job to size it correctly.
GENERATION_STEPS = [
    "Connecting to SolidWorks",
    "Opening GAD template",
    "Sheet 1: inserting GA block",
    "Sheet 1: inserting hookup schematic",
    "Sheet 1 complete - title block properties set",
    "Sheet 2 complete - reference drawing inserted",
    "Sheet 3 complete - nameplate inserted",
    "Rebuilding drawing",
    "Exporting PDF",
]
GENERATION_TOTAL_STEPS = len(GENERATION_STEPS)


def _noop_progress(step_text, step_index):
    pass


# See generate_from_resolved()'s docstring for why this exists.
_generation_lock = threading.Lock()

# SolidWorks' two automation type libraries - stable GUIDs across every
# SolidWorks version (only the registered (major, minor) under each changes
# release to release, read from the registry in _typelib_version() below).
#   - sldworks.tlb: the SldWorks.Application object model itself.
#   - swconst.tlb: the swXxx_e enums (swSaveAsCurrentVersion etc.) - a
#     separate type library, same split as the C# app's two "using"s
#     (SolidWorks.Interop.sldworks vs SolidWorks.Interop.swconst).
_SLDWORKS_TYPELIB_GUID = "{83A33D31-27C5-11CE-BFD4-00400513BB57}"
_SWCONST_TYPELIB_GUID = "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}"


class GadGenerationError(Exception):
    """User-facing failure - bad config, missing file, SolidWorks error,
    or (via GadLookupError) a lookup miss."""


class SolidWorksNotInstalledError(GadGenerationError):
    """Raised when this machine has no registered SolidWorks COM server."""


def solidworks_diagnostics() -> dict:
    """Reports whether SldWorks.Application resolves as a COM ProgID, plus
    which machine/Python process ran the check - surfaced via a diagnostic
    page so this can be inspected from a browser on whichever machine the
    Flask process is actually running on, without needing shell/remote
    access to it.

    Same technique as the original WinForms app's ConnectToSolidWorks()
    (Type.GetTypeFromProgID("SldWorks.Application") == null check): resolve
    the ProgID via HKEY_CLASSES_ROOT\\<ProgID>\\CLSID - the exact registry
    lookup the OS's ProgID-to-CLSID resolution (what .NET's
    Type.GetTypeFromProgID and, on the Python side, Dispatch()/
    CreateInstance() both use under the hood) performs. Doing this in the
    same process that will later call Dispatch() means it automatically
    sees the same 32-bit/64-bit registry view Dispatch() itself will -
    no manual WOW64 cross-checking needed, since both calls run under
    identical process bitness."""
    import platform
    import struct

    try:
        winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, "SldWorks.Application\\CLSID")
        installed = True
    except OSError:
        installed = False

    return {
        "installed": installed,
        "hostname": platform.node(),
        "python_bitness": struct.calcsize("P") * 8,
    }


def is_solidworks_installed() -> bool:
    return solidworks_diagnostics()["installed"]


def _config(key: str) -> str:
    value = os.environ.get(key)
    if not value:
        raise GadGenerationError(
            f"Missing required setting '{key}'. Set it in .env (see .env.example)."
        )
    return value


def _require_file(path: str, what: str):
    if not os.path.exists(path):
        raise GadGenerationError(f"{what} not found: {path}")


def _save_as_result(raw):
    """SaveAs has two [in,out] params (errors, warnings). Early-bound
    pywin32 dispatch returns them appended to the result as a tuple; fall
    back to treating the call as a plain bool if that's not how this
    particular COM binding surfaced it."""
    if isinstance(raw, tuple):
        return raw[0]
    return raw


def generate_globe_gad(row: dict, progress_cb=None, on_pdf_ready=None, close_event=None) -> str:
    """Runs the full Sheet1/2/3 pipeline for one Excel row and returns the
    exported PDF's path. Raises GadGenerationError / GadLookupError on any
    failure, same failure points as the WinForms version's MessageBox calls.

    Combines the database lookups and the SolidWorks automation in one call
    - only usable on a machine that has both a DB connection and SolidWorks
    (e.g. worker.py). For the "resolve on the server, generate locally with
    no DB dependency" split used by the downloadable GadGenerate.exe, see
    app.services.gad_globe_lookup.resolve_globe_gad() + generate_from_resolved()
    below instead.

    progress_cb, on_pdf_ready and close_event: see generate_from_resolved()."""
    series = cell(row, "Series")
    if series not in ("10", "11"):
        raise GadGenerationError("This page only generates Globe valve (Series 10/11) drawings.")

    resolved = lookup.resolve_globe_gad(row)
    return generate_from_resolved(
        resolved, progress_cb=progress_cb, on_pdf_ready=on_pdf_ready, close_event=close_event
    )


# How long to keep SolidWorks open awaiting a close_event signal before
# closing it anyway - a safety net for an abandoned browser tab (closed,
# crashed, or the user just never clicked anything) so this thread (and its
# open SolidWorks process) doesn't sit around forever.
_CLOSE_WAIT_TIMEOUT_SECONDS = 1200


def generate_from_resolved(resolved: dict, progress_cb=None, on_pdf_ready=None, close_event=None) -> str:
    """Runs the full Sheet1/2/3 SolidWorks pipeline from an already-resolved
    dict (see app.services.gad_globe_lookup.resolve_globe_gad) and returns
    the exported PDF's path. No database access here at all - only
    SolidWorks and the local GAD_* block/template files (see _config), so
    this is what actually runs inside GadGenerate.exe on a user's own
    machine, entirely offline apart from SolidWorks itself.

    progress_cb, if given, is called as progress_cb(step_text, step_index)
    after each milestone in GENERATION_STEPS completes - see
    app/blueprints/dashboard/routes.py's /globe/generate-server for how the
    browser polls this to show real step-by-step progress.

    on_pdf_ready, if given, is called as on_pdf_ready(pdf_path, sldworks_path)
    as soon as both files are saved - letting the caller mark the job "done"
    (e.g. so the browser can download them) *before* this function actually
    returns, since it then keeps running to handle close_event below.

    close_event, if given, is a threading.Event: once the PDF is ready, this
    blocks (instead of returning immediately) until that event is set (or
    _CLOSE_WAIT_TIMEOUT_SECONDS elapses), then closes SolidWorks - discarding
    any unsaved changes, no save prompt - before finally returning. This has
    to happen on this exact thread: SolidWorks's COM object is apartment-
    threaded, so only the thread that created it (this one) can safely call
    it again later, which is why "close SolidWorks after the user
    acknowledges" is a wait-then-continue here rather than a separate call
    from whatever thread handles that later HTTP request."""
    report = progress_cb or _noop_progress
    sw_app = None

    # SolidWorks is COM, and COM requires the calling thread to have an
    # initialized apartment before any Dispatch() call - true for a plain
    # script's main thread (worker.py) but NOT for a Flask request thread,
    # which raises "CoInitialize has not been called" otherwise.
    pythoncom.CoInitialize()
    try:
        # SolidWorks COM automation can only do one thing at a time - the
        # old synchronous route serialized concurrent clicks for free (one
        # Flask request blocked until done before the next started); this
        # lock re-establishes that guarantee explicitly for the actual
        # drawing-building steps. It's released before the close_event wait
        # below, so one job sitting open awaiting the user's acknowledgment
        # doesn't block everyone else's generations.
        with _generation_lock:
            sw_app = _connect()
            report(GENERATION_STEPS[0], 0)

            model, drawing = _open_gad_template(sw_app)
            report(GENERATION_STEPS[1], 1)

            _generate_sheet1(sw_app, model, drawing, resolved, report)
            # Sheet2 deletes and replaces the template's own Sheet2 tab (see
            # its docstring), which invalidates the model/drawing COM
            # references above - it returns fresh ones to use from here on.
            model, drawing = _generate_sheet2(sw_app, model, drawing, resolved, report)
            _generate_sheet3(sw_app, model, drawing, resolved, report)

            model.ForceRebuild3(False)
            report(GENERATION_STEPS[7], 7)

            # Defaults to the old app's behavior (export straight to Desktop)
            # if GAD_OUTPUT_DIR isn't set.
            output_dir = os.environ.get("GAD_OUTPUT_DIR") or os.path.join(
                os.path.expanduser("~"), "Desktop"
            )
            if not os.path.isdir(output_dir):
                output_dir = tempfile.gettempdir()
            os.makedirs(output_dir, exist_ok=True)
            sldworks_path = os.path.join(output_dir, "GA_Drawing.SLDDRW")
            pdf_path = os.path.join(output_dir, "GA_Drawing.pdf")

            # Save the native SolidWorks drawing first - SaveAs infers the
            # format from the extension, same call as the PDF export below
            # just pointed at a different path. Doing this before the PDF
            # export means the document is a saved file (not the untitled
            # "Draw1" NewDocument() created) for that second SaveAs too,
            # though that call doesn't actually depend on it.
            result = model.Extension.SaveAs(
                sldworks_path,
                sw_const.swSaveAsCurrentVersion,
                sw_const.swSaveAsOptions_Silent,
                None,
                0,
                0,
            )
            if not _save_as_result(result):
                raise GadGenerationError("SolidWorks drawing save failed.")

            result = model.Extension.SaveAs(
                pdf_path,
                sw_const.swSaveAsCurrentVersion,
                sw_const.swSaveAsOptions_Silent,
                None,
                0,
                0,
            )
            if not _save_as_result(result):
                raise GadGenerationError("PDF export failed.")
            report(GENERATION_STEPS[8], 8)

        if on_pdf_ready:
            on_pdf_ready(pdf_path, sldworks_path)

        if close_event is not None:
            # Leave SolidWorks open (Visible=True, same as the original) so
            # the user can inspect the drawing, until they acknowledge the
            # "PDF generated" message (or the safety timeout above fires).
            close_event.wait(timeout=_CLOSE_WAIT_TIMEOUT_SECONDS)
            try:
                sw_app.CloseAllDocuments(True)  # True = discard unsaved changes, no prompt
                sw_app.ExitApp()
            except Exception:
                pass  # best-effort - a user closing SolidWorks by hand first is fine too

        return pdf_path
    finally:
        pythoncom.CoUninitialize()


def _typelib_version(guid: str) -> tuple:
    """Reads whichever (major, minor) version of this type library is
    actually registered on this machine, from
    HKEY_CLASSES_ROOT\\TypeLib\\<guid>\\<major>.<minor> - rather than
    hardcoding a SolidWorks release year, so this keeps working across
    SolidWorks upgrades with no code change. If more than one version is
    registered (e.g. a stale entry left over from a previous SolidWorks
    install alongside the current one), the highest wins."""
    versions = []
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"TypeLib\\{guid}") as key:
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(key, i)
            except OSError:
                break
            i += 1
            try:
                major_str, minor_str = sub.split(".")
                versions.append((int(major_str, 16), int(minor_str, 16)))
            except ValueError:
                continue
    if not versions:
        raise GadGenerationError(f"No registered type library found for {guid}.")
    return max(versions)


def _sldworks_module():
    """Loads (or returns the already-cached) pywin32 generated wrapper
    module for SolidWorks' main type library, built directly from its
    registered version (see _typelib_version) - never through a live
    instance's GetTypeInfo(), which is what _wrap() below needs to re-bind
    dynamic return values early."""
    major, minor = _typelib_version(_SLDWORKS_TYPELIB_GUID)
    return win32com.client.gencache.EnsureModule(_SLDWORKS_TYPELIB_GUID, 0, major, minor)


def _ensure_typelibs():
    """Builds pywin32's generated wrapper modules for both SolidWorks type
    libraries directly from their registered typelib info (see
    _typelib_version), which is what makes win32com.client.constants.*
    (sw_const below) resolve, and what _wrap() uses to re-bind dynamic
    return values (see its docstring for why that's needed)."""
    _sldworks_module()
    major, minor = _typelib_version(_SWCONST_TYPELIB_GUID)
    win32com.client.gencache.EnsureModule(_SWCONST_TYPELIB_GUID, 0, major, minor)


def _wrap(dynamic_obj, classname):
    """Re-binds a dynamically-dispatched COM return value as the known
    early-bound SolidWorks interface class named by `classname`.

    Several SolidWorks methods - NewDocument(), GetMathUtility() - declare a
    plain IDispatch return type in the type library, same reason the C# port
    needs an explicit cast ("(DrawingDoc)swModel", "(MathUtility)swApp.
    GetMathUtility()"). Left as plain dynamic dispatch in Python, calling
    members on them (ActivateSheet, InsertSketch, ...) fails outright with
    "Member not found" - SolidWorks's IDispatch implementation doesn't
    resolve those names for a dynamically-dispatched object. Re-wrapping
    with the type library's own known interface class (looked up from the
    already-loaded module, not from the object's own GetTypeInfo() - see
    _connect()) fixes it, exactly mirroring the C# cast."""
    mod = _sldworks_module()
    target_class = getattr(mod, classname)
    target_class = getattr(target_class, "default_interface", target_class)
    return target_class(dynamic_obj._oleobj_)


def _connect():
    if not is_solidworks_installed():
        import platform

        raise SolidWorksNotInstalledError(
            f"This computer ({platform.node()}) does not have SolidWorks "
            "installed. GAD generation only works on a machine that has "
            "SolidWorks installed and licensed."
        )
    try:
        _ensure_typelibs()
        # Deliberately NOT win32com.client.gencache.EnsureDispatch("SldWorks.
        # Application"): EnsureDispatch identifies an object's type library
        # by calling GetTypeInfo() on the live instance, and SolidWorks's
        # SldWorks.Application fails that specific call with "Element not
        # found" (pythoncom.com_error -2147319765) even though the object
        # itself works fine - which is what surfaces as pywin32's generic
        # "This COM object can not automate the makepy process" error.
        # Instead: get the early-bound class straight from the ProgID's
        # registered CLSID (no live-instance call involved), and construct
        # it from a plain Dispatch()'s raw COM pointer.
        klass = win32com.client.gencache.GetClassForProgID("SldWorks.Application")
        raw = win32com.client.Dispatch("SldWorks.Application")
        sw_app = klass(raw._oleobj_)
    except Exception as exc:
        raise GadGenerationError(f"Could not start SolidWorks: {exc}") from exc
    sw_app.Visible = True
    return sw_app


def _open_gad_template(sw_app):
    template_path = _config("GAD_TEMPLATE_PATH")
    _require_file(template_path, "Template")

    model_raw = sw_app.NewDocument(template_path, sw_const.swDwgPapersUserDefined, 0, 0)
    if model_raw is None:
        raise GadGenerationError(f"Failed to create new drawing from template: {template_path}")
    # See _wrap()'s docstring - NewDocument's return needs re-binding to
    # both interfaces we use, same as the C# port's separate ModelDoc2/
    # DrawingDoc-typed variables for the same underlying document.
    model = _wrap(model_raw, "ModelDoc2")
    drawing = _wrap(model_raw, "DrawingDoc")
    return model, drawing


def _insert_block(sw_app, sketch_mgr, math_util, xyz, block_path, what):
    _require_file(block_path, what)
    point = math_util.CreatePoint(list(xyz))
    sketch_mgr.MakeSketchBlockFromFile(point, block_path, True, BLOCK_INSERT_SCALE, 0)


def _generate_sheet1(sw_app, model, drawing, resolved, report):
    drawing.ActivateSheet("Sheet1")
    sketch_mgr = model.SketchManager
    math_util = _wrap(sw_app.GetMathUtility(), "MathUtility")
    sketch_mgr.InsertSketch(True)

    # --- GA block ---
    block_dir = _config("GAD_BLOCK_DIR")
    ga_block = os.path.join(block_dir, f"{resolved['drawing_no']}.SLDBLK")
    _insert_block(sw_app, sketch_mgr, math_util, (0.32, 0.52, 10), ga_block, "GA block")
    report(GENERATION_STEPS[2], 2)

    # --- Hookup schematic block ---
    hookup_dir = _config("GAD_HOOKUP_BLOCK_DIR")
    hookup_block = os.path.join(hookup_dir, f"{resolved['hookup_no']}.SLDBLK")
    _insert_block(sw_app, sketch_mgr, math_util, (0.0, 0.0, 0), hookup_block, "Hookup schematic block")
    report(GENERATION_STEPS[3], 3)

    sketch_mgr.InsertSketch(True)

    # --- Custom properties for the title block (already fully resolved -
    # both the plain pass-through fields and the A/B/C/AR/D/E/F/WEIGHT
    # dimension values live in the same flat dict) ---
    prop_mgr = model.Extension.CustomPropertyManager("")
    for name, value in resolved["properties"].items():
        prop_mgr.Set2(name, value)
    report(GENERATION_STEPS[4], 4)


def _generate_sheet2(sw_app, model, drawing, resolved, report):
    """Replaces the template's own Sheet2 tab outright with a whole separate
    .SLDDRW reference drawing (GAD_SHEET4_DIR/<resolved['sheet4_no']>.SLDDRW,
    looked up via app.services.gad_globe_lookup.find_sheet4_drawing_no) -
    Sheet2 no longer holds a .SLDBLK cross-section block, so this can't use
    _insert_block()/MakeSketchBlockFromFile like Sheet1/Sheet3. SolidWorks
    has no "insert sheet from another file" API call, so this mirrors the
    UI's own "copy sheet, paste sheet" feature instead, verified reliably
    (multiple repeated runs, including with real Sheet1 content already in
    place) against a live SolidWorks 2022 session:
      1. Delete the template's existing Sheet2 (selected as a whole SHEET
         entity, not its contents, via Extension.DeleteSelection2) - this
         doesn't prompt for confirmation the way the UI's Delete key would.
      2. Open the reference file as its own document.
      3. Select its (only) sheet as a whole SHEET entity and EditCopy it -
         this puts the whole sheet on the clipboard.
      4. Switch back to our drawing, select Sheet1 (the paste needs some
         sheet selected to anchor to), and PasteSheet. IMPORTANT: the
         source document from step 2 must stay open until *after* this
         call - closing it first (even right after EditCopy) makes
         PasteSheet silently return False every time. SolidWorks' "sheet on
         the clipboard" isn't a plain OS clipboard payload; it still
         depends on the source document being alive.
      5. The pasted sheet lands with the source's own sheet name (e.g.
         "Sheet2", from the reference file's own internal sheet name) -
         renamed to "Sheet2" via ISheet.SetName (a no-op if it already
         happened to match, otherwise resolving pywin32's own "(2)"
         auto-rename of the collision). It also always lands right after
         Sheet1 regardless of the anchor above (a real quirk of PasteSheet
         on this SolidWorks version, not a mistake here) - fixed explicitly
         with ReorderSheets afterward, using a SAFEARRAY of BSTR (a plain
         Python list marshals as the wrong VARIANT subtype and ReorderSheets
         just silently returns False instead of reordering).
      6. Only now close the source document.

    Deleting a sheet disconnects pywin32's cached model/drawing COM
    references from the live document (later calls on them raise "The
    object invoked has disconnected from its clients") - this re-wraps
    fresh ones from sw_app.ActiveDoc before returning, which the caller
    must use from here on instead of its own now-stale model/drawing."""
    sheet4_dir = _config("GAD_SHEET4_DIR")
    sheet2_path = os.path.join(sheet4_dir, f"{resolved['sheet4_no']}.SLDDRW")
    _require_file(sheet2_path, "Sheet2 reference drawing")

    model.Extension.SelectByID2("Sheet2", "SHEET", 0, 0, 0, False, 0, None, 0)
    model.Extension.DeleteSelection2(0)

    result = sw_app.OpenDoc6(
        sheet2_path, sw_const.swDocDRAWING, sw_const.swOpenDocOptions_Silent, "", 0, 0
    )
    src_raw = result[0] if isinstance(result, tuple) else result
    if src_raw is None:
        raise GadGenerationError(f"Failed to open Sheet2 reference drawing: {sheet2_path}")
    src_model = _wrap(src_raw, "ModelDoc2")
    src_drawing = _wrap(src_raw, "DrawingDoc")

    try:
        src_sheet_name = src_drawing.GetSheetNames()[0]
        src_model.Extension.SelectByID2(src_sheet_name, "SHEET", 0, 0, 0, False, 0, None, 0)
        src_model.EditCopy()

        sw_app.ActivateDoc2(model.GetTitle(), False, 0)
        model.Extension.SelectByID2("Sheet1", "SHEET", 0, 0, 0, False, 0, None, 0)
        # Checked explicitly because a failed paste doesn't raise on its
        # own, and an earlier version of this function that skipped the
        # check went on to rename whatever GetCurrentSheet() returned when
        # no new sheet was actually created - silently corrupting the
        # drawing instead of failing loudly.
        if not drawing.PasteSheet(sw_const.swInsertOption_AfterSelectedSheet, sw_const.swRenameOption_No):
            raise GadGenerationError("Could not paste the Sheet2 reference drawing's sheet.")

        new_sheet = _wrap(drawing.GetCurrentSheet(), "Sheet")
        new_sheet.SetName("Sheet2")
    finally:
        sw_app.CloseDoc(src_model.GetTitle())

    ordered = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_BSTR, ["Sheet1", "Sheet2", "Sheet3"]
    )
    if not drawing.ReorderSheets(ordered):
        raise GadGenerationError("Could not order Sheet2 after Sheet1.")

    report(GENERATION_STEPS[5], 5)

    fresh_raw = sw_app.ActiveDoc
    return _wrap(fresh_raw, "ModelDoc2"), _wrap(fresh_raw, "DrawingDoc")


def _generate_sheet3(sw_app, model, drawing, resolved, report):
    drawing.ActivateSheet("Sheet3")
    sketch_mgr = model.SketchManager
    math_util = _wrap(sw_app.GetMathUtility(), "MathUtility")
    sketch_mgr.InsertSketch(True)

    # TODO: same placeholder as the original app - this is a fixed
    # nameplate block, not looked up per-row (resolved["nameplate"] carries
    # the same fixed placeholder for now - see resolve_globe_gad). Confirm
    # the real nameplate source (a DB lookup like Sheet1/2, or a fixed file
    # per valve type) and wire it in here instead.
    nameplate_block = _config("GAD_NAMEPLATE_BLOCK_PATH")
    _insert_block(sw_app, sketch_mgr, math_util, (0.0, 0.0, 0), nameplate_block, "Nameplate block")

    sketch_mgr.InsertSketch(True)
    report(GENERATION_STEPS[6], 6)


