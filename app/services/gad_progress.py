"""In-memory progress tracking for the "Generate on Server" button, so the
browser can show real step-by-step status (see templates/globe.html) while
SolidWorks runs, instead of one long spinner.

Deliberately NOT the GadGenerationJob database table - that's the separate
worker.py queue for the "download a package, run it on your own machine"
flow, meant to survive across processes/machines. This is transient,
single-process, in-memory state for one synchronous "generate right now on
this server" click; it's fine for it to vanish on a server restart, and
keeping it out of the database means nothing here needs a migration.
"""
import threading
import time
import uuid

_lock = threading.Lock()
_jobs = {}  # token -> dict

# Finished (or abandoned - e.g. the tab was closed mid-poll) jobs are swept
# after this long so _jobs can't grow forever on a long-running process.
# Kept comfortably longer than solidworks_automation's
# _CLOSE_WAIT_TIMEOUT_SECONDS so a job's close_event (see start_job) stays
# reachable via request_close() for the whole time its generation thread
# could still be waiting on it.
_TTL_SECONDS = 1800


def _sweep_locked():
    cutoff = time.time() - _TTL_SECONDS
    stale = [token for token, job in _jobs.items() if job["created_at"] < cutoff]
    for token in stale:
        _jobs.pop(token, None)


def start_job(total_steps: int) -> str:
    """Registers a new job and returns its token - hand this back to the
    browser so it can poll get_job(token) for progress.

    step_index tracks the index of the last *completed* step (see update()),
    so it starts at -1 here - nothing has completed yet, step 0 is the one
    about to run. The frontend renders step (step_index + 1) as the
    currently active one.

    Also creates a threading.Event (see get_close_event/request_close) that
    solidworks_automation.py's generation thread waits on before closing
    SolidWorks, once the PDF is ready - lets the browser tell that same
    thread "the user clicked OK, close it now" without ever calling
    SolidWorks's COM object from a different thread (see that module's
    docstring for why that matters)."""
    token = uuid.uuid4().hex
    with _lock:
        _sweep_locked()
        _jobs[token] = {
            "step": "Starting...",
            "step_index": -1,
            "total_steps": total_steps,
            "done": False,
            "error": None,
            "error_table_name": None,
            "error_table_label": None,
            "pdf_path": None,
            "download_name": None,
            "sldworks_path": None,
            "sldworks_download_name": None,
            "created_at": time.time(),
            "close_event": threading.Event(),
        }
    return token


def get_close_event(token: str):
    """Returns the job's close_event (see start_job) for the generation
    thread to wait on, or None if the token is unknown/expired."""
    with _lock:
        job = _jobs.get(token)
        return job["close_event"] if job is not None else None


def request_close(token: str) -> bool:
    """Signals the job's generation thread to close SolidWorks now (it may
    already have, via the safety timeout). Returns False if the token is
    unknown/expired - e.g. it already got swept - in which case SolidWorks
    for that job will only close via that timeout."""
    with _lock:
        job = _jobs.get(token)
        if job is None:
            return False
        job["close_event"].set()
        return True


def update(token: str, step: str, step_index: int):
    with _lock:
        job = _jobs.get(token)
        if job is not None:
            job["step"] = step
            job["step_index"] = step_index


def finish(
    token: str,
    pdf_path: str,
    download_name: str,
    sldworks_path: str = None,
    sldworks_download_name: str = None,
):
    with _lock:
        job = _jobs.get(token)
        if job is not None:
            job["done"] = True
            job["step"] = "Done"
            job["step_index"] = job["total_steps"]
            job["pdf_path"] = pdf_path
            job["download_name"] = download_name
            job["sldworks_path"] = sldworks_path
            job["sldworks_download_name"] = sldworks_download_name


def fail(token: str, error: str, table_name: str = None, table_label: str = None):
    """table_name/table_label, when the failure was a GadLookupError, name
    the specific GAD Masters table that's missing a matching row - see that
    exception's docstring - so the browser can link straight to it instead
    of leaving the user to guess which of the 6 master tables to check."""
    with _lock:
        job = _jobs.get(token)
        if job is not None:
            job["done"] = True
            job["error"] = error
            job["error_table_name"] = table_name
            job["error_table_label"] = table_label


def get_job(token: str):
    with _lock:
        job = _jobs.get(token)
        return dict(job) if job is not None else None
