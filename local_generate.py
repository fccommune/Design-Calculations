"""
Local, no-database GAD drawing generator - packaged as GadGenerate.exe.

The web app resolves which master-table rows match a BOM row (GA drawing
number, hookup schematic, cross-section, dimension values - see
app.services.gad_globe_lookup.resolve_globe_gad) *before* this ever runs,
when the "Generate GAD" button is clicked, and hands the browser a small
zip containing this program plus that resolved data as job.json.

Double-clicking GadGenerate.exe reads job.json from its own folder and
drives SolidWorks using only those already-resolved values - see
app.services.solidworks_automation.generate_from_resolved(). No database,
no network access, needed here at all: only requirements on this machine
are SolidWorks itself (installed and licensed) and a .env file next to
this exe pointing GAD_TEMPLATE_PATH / GAD_BLOCK_DIR / etc. at this
machine's own copy of the SolidWorks Automation File folder.

Usage:
    GadGenerate.exe                    (reads job.json from its own folder)
    GadGenerate.exe path\\to\\job.json   (reads a specific file instead)
"""
import json
import os
import sys

_BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))

from dotenv import load_dotenv

load_dotenv(os.path.join(_BASE_DIR, ".env"))

if getattr(sys, "frozen", False):
    # win32com.client.gencache.EnsureDispatch (used in
    # solidworks_automation.py to talk to SolidWorks) generates/caches a
    # Python wrapper for SolidWorks' COM type library on first use, normally
    # under the pywin32 install's own folder - which doesn't exist in a
    # frozen exe. Point it at a writable per-machine temp folder instead, or
    # this fails the first time SolidWorks is dispatched.
    import tempfile
    import win32com
    win32com.__gen_path__ = os.path.join(tempfile.gettempdir(), "gen_py")

from app.services.solidworks_automation import (
    generate_from_resolved,
    GadGenerationError,
    SolidWorksNotInstalledError,
)


def main():
    job_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(_BASE_DIR, "job.json")

    print("=" * 60)
    print("GAD Drawing Generator")
    print("=" * 60)

    if not os.path.exists(job_path):
        print(f"\njob.json not found: {job_path}")
        print("Make sure job.json is in the same folder as this program")
        print("(it comes bundled together when you download the package).")
        input("\nPress Enter to close...")
        sys.exit(1)

    env_path = os.path.join(_BASE_DIR, ".env")
    if not os.path.exists(env_path):
        print(f"\n.env not found: {env_path}")
        print("Create one next to this exe with this machine's own paths -")
        print("see .env.example in this same folder.")
        input("\nPress Enter to close...")
        sys.exit(1)

    with open(job_path, "r", encoding="utf-8") as f:
        resolved = json.load(f)

    dwg_no = resolved.get("properties", {}).get("DWG_NO") or "this valve"
    print(f"\nGenerating GAD drawing for {dwg_no}...")
    print("Opening SolidWorks - this can take a minute. Please wait.\n")

    try:
        pdf_path = generate_from_resolved(resolved)
    except SolidWorksNotInstalledError as exc:
        print(f"SolidWorks not found: {exc}")
        input("\nPress Enter to close...")
        sys.exit(1)
    except GadGenerationError as exc:
        print(f"Generation failed: {exc}")
        input("\nPress Enter to close...")
        sys.exit(1)
    except Exception as exc:  # unexpected SolidWorks/COM failure
        print(f"Unexpected error: {exc}")
        input("\nPress Enter to close...")
        sys.exit(1)

    print(f"Done! PDF saved to:\n{pdf_path}")
    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
