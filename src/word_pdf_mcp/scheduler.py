from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path

from .errors import ConversionError
from .files import (
    atomic_json,
    inspect_package,
    publish_pdf,
    sha256,
    validate_paths,
    validate_pdf,
)
from .mutex import MUTEX_NAME, word_mutex
from .processes import identity, matching_process, terminate_owned_word, wait_gone


def state_directory() -> Path:
    if sys.platform != "win32":
        raise ConversionError("windows_required", "Run this MCP on Windows with desktop Word.")
    import win32ts

    # The journal has the same session scope as the Local\\ mutex.
    session = win32ts.ProcessIdToSessionId(os.getpid())
    return Path(os.environ["LOCALAPPDATA"]) / "WordPdfMcp" / f"session-{session}"


def read_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ConversionError(
            "recovery_required", f"Cannot read recovery record {path}: {exc}"
        ) from exc


def cleanup_worker(pending: dict) -> None:
    """Run only while holding the mutex, including when it was abandoned.

    Unknown Word ownership intentionally blocks new conversions instead of guessing
    which of the user's Word processes is safe to terminate.
    """
    task_dir = Path(pending["task_dir"])
    state_path = task_dir / "state.json"
    # Stop a child left behind by a dead supervisor before assessing Word ownership.
    worker_record = pending.get("worker")
    if worker_record:
        process = matching_process(worker_record)
        if process is not None:
            process.kill()
        if not wait_gone(worker_record):
            raise ConversionError(
                "recovery_required", "Previous conversion worker is still running."
            )
    if not state_path.exists():
        if worker_record:
            return  # Worker publishes state before attempting to launch Word.
        # Crash between Popen and recording child identity: do not assume no child exists.
        raise ConversionError(
            "recovery_required", "Worker startup ownership is unknown; inspect state."
        )
    state = read_json(state_path)
    # The worker record can exist here even if the supervisor died before recording it.
    if state.get("worker"):
        process = matching_process(state["worker"])
        if process is not None:
            process.kill()
        if not wait_gone(state["worker"]):
            raise ConversionError("recovery_required", "Previous worker could not be stopped.")
    if state.get("word"):
        terminate_owned_word(state["word"])
    elif state.get("phase") != "before_word":
        raise ConversionError(
            "recovery_required",
            "Word may have started before its PID was recorded. Close the leftover automation "
            f"Word manually, then follow docs/RECOVERY.md. Task directory: {task_dir}",
        )


class ConversionService:
    def __init__(self, root: Path | None = None):
        self._root = root
        self._local_lock = threading.Lock()

    @property
    def root(self) -> Path:
        return self._root if self._root is not None else state_directory()

    def health(self) -> dict:
        if sys.platform != "win32":
            return {"platform": sys.platform, "ready": False, "reason": "Windows required"}
        root = self.root
        return {
            "platform": "win32",
            "mutex": MUTEX_NAME,
            "state_directory": str(root),
            "pending_record": str(root / "pending.json")
            if (root / "pending.json").exists()
            else None,
            "note": "Pending may mean a running task; Word and fidelity are not probed.",
        }

    def convert(
        self,
        input_path: str,
        output_path: str,
        *,
        overwrite: bool = False,
        include_markup: bool = False,
        timeout_seconds: int = 180,
        queue_timeout_seconds: int = 600,
    ) -> dict:
        if sys.platform != "win32":
            raise ConversionError("windows_required", "Native Word conversion is Windows-only.")
        if not 10 <= timeout_seconds <= 1800 or not 1 <= queue_timeout_seconds <= 3600:
            raise ConversionError("invalid_timeout", "Conversion: 10–1800s; queue: 1–3600s.")
        started = time.monotonic()
        if not self._local_lock.acquire(timeout=queue_timeout_seconds):
            raise ConversionError("queue_timeout", "Timed out waiting for this MCP's worker.")
        try:
            remaining = queue_timeout_seconds - (time.monotonic() - started)
            if remaining <= 0:
                raise ConversionError("queue_timeout", "Queue deadline exceeded.")
            with word_mutex(remaining) as abandoned:
                root = self.root
                root.mkdir(parents=True, exist_ok=True)
                journal = root / "pending.json"
                if journal.exists():
                    cleanup_worker(read_json(journal))
                    journal.unlink()
                # Validate again under the global lock, so competing output names are safe.
                src, dst = validate_paths(input_path, output_path, overwrite)
                return self._run(
                    src,
                    dst,
                    journal,
                    overwrite,
                    include_markup,
                    timeout_seconds,
                    abandoned=abandoned,
                )
        finally:
            self._local_lock.release()

    def _run(
        self,
        src: Path,
        dst: Path,
        journal: Path,
        overwrite: bool,
        include_markup: bool,
        timeout: int,
        *,
        abandoned: bool,
    ) -> dict:
        task_dir = Path(tempfile.mkdtemp(prefix="task-", dir=self.root))
        # Staging in the destination filesystem makes final publication atomic.
        stage = dst.parent / f".wordpdf-{task_dir.name}.pdf"
        pending = {
            "task_dir": str(task_dir),
            "stage": str(stage),
            "supervisor": identity(os.getpid()),
        }
        process = None
        clean = False
        journal_written = False
        try:
            original_hash = sha256(src)
            copied = task_dir / ("source" + src.suffix.lower())
            shutil.copyfile(src, copied)
            if sha256(copied) != original_hash:
                raise ConversionError("source_changed", "Source changed while copying; retry.")
            inventory = inspect_package(copied)
            request = {
                "input_path": str(copied),
                "output_path": str(stage),
                "include_markup": include_markup,
                "parent": pending["supervisor"],
            }
            atomic_json(task_dir / "request.json", request)
            atomic_json(journal, pending)
            journal_written = True
            with (task_dir / "worker.log").open("wb") as log:
                try:
                    process = subprocess.Popen(
                        [
                            sys.executable,
                            "-m",
                            "word_pdf_mcp.worker",
                            str(task_dir / "request.json"),
                            str(task_dir / "state.json"),
                            str(task_dir / "result.json"),
                        ],
                        stdin=subprocess.DEVNULL,
                        stdout=log,
                        stderr=log,
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )
                except OSError as exc:
                    clean = True  # Popen did not create a worker, so no Word could be launched.
                    raise ConversionError("worker_start_failed", str(exc)) from exc
                pending["worker"] = identity(process.pid)
                atomic_json(journal, pending)
                try:
                    process.wait(timeout=timeout)
                except subprocess.TimeoutExpired as exc:
                    process.kill()
                    process.wait(timeout=10)
                    raise ConversionError(
                        "conversion_timeout", f"Word conversion exceeded {timeout}s."
                    ) from exc
            result = read_json(task_dir / "result.json")
            if process.returncode != 0 or result.get("status") != "success":
                raise ConversionError("word_failed", result.get("message", "Worker failed."))
            cleanup_worker(pending)
            clean = True
            pdf = validate_pdf(stage)
            if pdf["pages"] != result["word_pages"]:
                raise ConversionError(
                    "page_count_mismatch",
                    f"Word reports {result['word_pages']} pages; PDF contains {pdf['pages']}.",
                )
            if sha256(src) != original_hash:
                raise ConversionError("source_changed", "Source changed during conversion; retry.")
            publish_pdf(stage, dst, overwrite)
            return {
                **result,
                **pdf,
                "output_path": str(dst),
                "input_sha256": original_hash,
                "output_sha256": sha256(dst),
                "inventory": inventory,
                "recovered_abandoned_mutex": abandoned,
                "visual_fidelity_verified": False,
            }
        finally:
            # Keep the lock throughout cleanup. Never allow the next task to overlap a survivor.
            try:
                if process is not None and process.poll() is None:
                    process.kill()
                    process.wait(timeout=10)
                if journal_written and not clean:
                    cleanup_worker(pending)
                    clean = True
            except Exception as exc:
                # Retain journal + task data; the next caller must verify recovery first.
                raise ConversionError(
                    "recovery_required", f"Cleanup could not be verified: {exc}. State: {task_dir}"
                ) from exc
            finally:
                if clean or not journal_written:
                    if journal_written:
                        journal.unlink(missing_ok=True)
                    stage.unlink(missing_ok=True)
                    shutil.rmtree(task_dir)
