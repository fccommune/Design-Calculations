"""
Background worker for GAD generation jobs.

Polls gad_generation_job for pending rows and drives SolidWorks on
whichever machine this script runs on - run it only on a machine that has
SolidWorks installed and licensed (see app/services/solidworks_automation.py,
unchanged from the synchronous version this replaces).

This decouples the web app (login, BOM upload, GAD Masters, the row
editor's dropdown/mismatch validation - none of that needs SolidWorks) from
the actual "click Generate GAD -> SolidWorks opens and exports a PDF" step:
clicking Generate just writes a 'pending' row to gad_generation_job and the
browser polls for it to flip to 'done'/'error' (see /globe/generate,
/globe/generate/<id>/status and /globe/generate/<id>/download in
app/blueprints/dashboard/routes.py). The web app and this worker only need
to share one database - they don't need to run on the same machine, and you
can run more than one of this worker (on different SolidWorks-licensed
machines) against the same queue; jobs are claimed with SELECT ... FOR
UPDATE SKIP LOCKED so they won't double-process a job.

Usage:
    python worker.py
    (or, packaged as a standalone exe - see build_worker_exe.md - just
    double-click GadWorker.exe; no Python install needed on that machine)

Runs until stopped with Ctrl+C.
"""
import msvcrt
import os
import platform
import sys
import time
from datetime import datetime

# When packaged by PyInstaller (sys.frozen), __file__-relative paths point
# into the temp extraction folder, not the machine the .exe actually lives
# on - resolve both the .env location and, below, the pywin32 COM wrapper
# cache off sys.executable's real directory instead.
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

from sqlalchemy import text

from app import create_app
from app.extensions import db
from app.models import GadGenerationJob
from app.services.gad_globe_lookup import cell, GadLookupError
from app.services.solidworks_automation import generate_globe_gad, GadGenerationError

POLL_SECONDS = 2
HOSTNAME = platform.node()

_CLAIM_SQL = text(
    """
    UPDATE gad_generation_job
    SET status = 'processing', started_at = :now, worker_hostname = :host
    WHERE id = (
        SELECT id FROM gad_generation_job
        WHERE status = 'pending'
        ORDER BY created_at
        FOR UPDATE SKIP LOCKED
        LIMIT 1
    )
    RETURNING id
    """
)


def _claim_next_job():
    row = db.session.execute(_CLAIM_SQL, {"now": datetime.utcnow(), "host": HOSTNAME}).fetchone()
    db.session.commit()
    if row is None:
        return None
    return db.session.get(GadGenerationJob, row[0])


def _process(job):
    row = job.row_data
    try:
        pdf_path = generate_globe_gad(row)
        with open(pdf_path, "rb") as f:
            job.pdf_data = f.read()
        dwg_no = cell(row, "Dwg No") or "GA_Drawing"
        job.download_name = f"{dwg_no}.pdf".replace("/", "-")
        job.status = GadGenerationJob.STATUS_DONE
    except (GadLookupError, GadGenerationError) as exc:
        job.status = GadGenerationJob.STATUS_ERROR
        job.error_message = str(exc)
    except Exception as exc:  # unexpected SolidWorks/COM failure
        job.status = GadGenerationJob.STATUS_ERROR
        job.error_message = f"GAD generation failed: {exc}"
    finally:
        job.finished_at = datetime.utcnow()
        db.session.commit()


_LOCK_PATH = os.path.join(_BASE_DIR, ".worker.lock")


def _acquire_single_instance_lock():
    """Refuses to start a second worker on this same machine - two workers
    here would each claim a different job (SELECT ... FOR UPDATE SKIP
    LOCKED only stops them claiming the *same* one) and drive two SolidWorks
    COM sessions at once, breaking the one-job-fully-finishes-before-the-
    next-starts guarantee. Held via an OS-level file lock (msvcrt.locking)
    that's released automatically if this process dies, so a crash can't
    leave a stale lock blocking the next run - unlike a plain lock file."""
    lock_file = open(_LOCK_PATH, "w")
    try:
        msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        print(
            "Another worker.py is already running on this machine. Refusing to "
            "start a second one - two would drive SolidWorks at the same time "
            "instead of one job finishing before the next starts. Stop the "
            "other instance first if you need to restart this one."
        )
        sys.exit(1)
    return lock_file  # keep this open for the life of the process - closing it releases the lock


def main():
    _acquire_single_instance_lock()
    app = create_app("development")
    print(f"GAD worker starting on {HOSTNAME} - polling every {POLL_SECONDS}s. Ctrl+C to stop.", flush=True)
    with app.app_context():
        while True:
            job = _claim_next_job()
            if job is None:
                time.sleep(POLL_SECONDS)
                continue
            print(f"Processing job #{job.id}...", flush=True)
            _process(job)
            print(f"Job #{job.id} -> {job.status}" + (f": {job.error_message}" if job.error_message else ""), flush=True)


if __name__ == "__main__":
    main()
