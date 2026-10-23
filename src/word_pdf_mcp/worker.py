"""One task per process. COM objects never leave this STA thread.

The supervisor owns the mutex. This process owns a kill-on-close Windows Job
containing only a positively identified, newly created Word process.
"""

from __future__ import annotations

import gc
import json
import os
import sys
import threading
import time
import uuid
from pathlib import Path

import psutil

from .files import atomic_json
from .processes import identity, matching_process


def word_window_pid(word) -> int:
    """Identify an empty Word instance without creating a document.

    Word.Application has no Hwnd property. Temporarily tag its hidden OpusApp
    window, resolve that window's PID, and restore the caption even on failure.
    No options or user documents are changed during this handshake.
    """
    import win32gui
    import win32process

    if int(word.Documents.Count) != 0:
        raise RuntimeError("Word already contains documents; refusing to use it")
    original_caption = word.Caption
    marker = "WordPdfMcp-" + uuid.uuid4().hex
    pids = set()

    def inspect(hwnd, _):
        if win32gui.GetClassName(hwnd) == "OpusApp":
            if marker in win32gui.GetWindowText(hwnd):
                pids.add(win32process.GetWindowThreadProcessId(hwnd)[1])

    try:
        word.Caption = marker
        win32gui.EnumWindows(inspect, None)
    finally:
        word.Caption = original_caption
    if len(pids) != 1:
        raise RuntimeError("Cannot uniquely identify Word's window process; refusing to use it")
    return pids.pop()


def export_pdf(document, output: str, include_markup: bool) -> str:
    """Fallback only for an absent API; an actual export failure must propagate."""
    options = {
        "OutputFileName": output,
        "ExportFormat": 17,
        "OpenAfterExport": False,
        "OptimizeFor": 0,
        "Range": 0,
        "Item": 7 if include_markup else 0,
        "IncludeDocProps": False,
        "KeepIRM": True,
        "CreateBookmarks": 0,
        "DocStructureTags": True,
        "BitmapMissingFonts": True,
        "UseISO19005_1": False,
    }
    try:
        method = document.ExportAsFixedFormat2
    except AttributeError:
        method = None
    if method is not None:
        method(**options, OptimizeForImageQuality=True)
        return "ExportAsFixedFormat2"
    document.ExportAsFixedFormat(**options)
    return "ExportAsFixedFormat"


def watch_parent(parent: dict) -> None:
    """Parent death must also close this process's Job handle and stop owned Word."""
    while True:
        try:
            if matching_process(parent) is None:
                os._exit(70)
        except Exception:
            os._exit(70)
        time.sleep(0.25)


def convert(request: dict, state_path: Path) -> dict:
    import pythoncom
    import pywintypes
    import win32api
    import win32con
    import win32event
    import win32job
    from win32com.client import DispatchEx

    state = {"phase": "before_word", "worker": identity(os.getpid())}
    atomic_json(state_path, state)
    threading.Thread(target=watch_parent, args=(request["parent"],), daemon=True).start()
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    word = document = job = word_handle = None
    owned = False
    saved_options = {}
    try:
        # Snapshot existing Word PIDs; creation time and an empty document set are also checked.
        before = {
            process.pid
            for process in psutil.process_iter(["name"])
            if (process.info["name"] or "").lower() == "winword.exe"
        }
        state["phase"] = "starting_word"
        atomic_json(state_path, state)
        launch_started = time.time()
        try:
            word = DispatchEx("Word.Application")
        except pywintypes.com_error as exc:
            # These registration failures cannot have launched the Word COM class.
            if exc.hresult in {-2147221164, -2147221005}:
                state["phase"] = "before_word"
                atomic_json(state_path, state)
            raise
        pid = word_window_pid(word)
        word_identity = identity(pid)
        if (
            pid in before
            or word_identity["name"].lower() != "winword.exe"
            or word_identity["created"] < launch_started - 0.1
            or int(word.Documents.Count) != 0
        ):
            # Do not call Quit or change settings on an instance we cannot own.
            raise RuntimeError("Word did not provide a verifiably new process; refusing to use it")
        owned = True
        state.update(phase="word_identified", word=word_identity)
        atomic_json(state_path, state)

        # pywin32 requires a string here; an empty name creates an unnamed job.
        job = win32job.CreateJobObject(None, "")
        limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        limits["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
        word_handle = win32api.OpenProcess(
            win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
            False,
            pid,
        )
        win32job.AssignProcessToJobObject(job, word_handle)
        state["phase"] = "job_assigned"
        atomic_json(state_path, state)

        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        for option in (
            "UpdateLinksAtOpen",
            "UpdateFieldsAtPrint",
            "UpdateLinksAtPrint",
            "SaveNormalPrompt",
        ):
            saved_options[option] = getattr(word.Options, option)
            setattr(word.Options, option, False)
        # Temporary copy: no library re-save, no compatibility upgrade, no object refresh.
        document = word.Documents.Open(
            FileName=request["input_path"],
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )
        document.Repaginate()
        pages = int(document.ComputeStatistics(2))  # wdStatisticPages
        api = export_pdf(document, request["output_path"], request["include_markup"])
        result = {
            "status": "success",
            "api": api,
            "word_pages": pages,
            "word_version": str(word.Version),
            "word_build": str(word.Build),
            "warnings": ["Fonts and equation/object appearance require visual acceptance testing."],
        }
        if api != "ExportAsFixedFormat2":
            result["warnings"].append(
                "Word lacks ExportAsFixedFormat2; original-image-quality option was unavailable."
            )
        return result
    finally:
        # Any hang here is bounded by the supervisor's whole-process timeout.
        try:
            if document is not None:
                document.Close(SaveChanges=0)
        finally:
            document = None
            try:
                if word is not None and owned:
                    for option, value in saved_options.items():
                        setattr(word.Options, option, value)
                    word.Quit(SaveChanges=0)
            finally:
                word = None
                gc.collect()
                # Closing the job also removes a leftover owned Word after Quit.
                if job is not None:
                    win32api.CloseHandle(job)
                if word_handle is not None:
                    win32event.WaitForSingleObject(word_handle, 5000)
                    win32api.CloseHandle(word_handle)
                pythoncom.CoUninitialize()


def main() -> int:
    request_path, state_path, result_path = map(Path, sys.argv[1:])
    try:
        request = json.loads(request_path.read_text(encoding="utf-8"))
        result = convert(request, state_path)
    except BaseException as exc:
        result = {"status": "error", "error": type(exc).__name__, "message": str(exc)}
    atomic_json(result_path, result)
    return 0 if result["status"] == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
