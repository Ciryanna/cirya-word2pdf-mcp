import json
import os
import subprocess
import sys
import time

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows kernel APIs required")


@pytest.mark.parametrize("enumeration_fails", [False, True])
def test_word_identity_failure_restores_caption(monkeypatch, enumeration_fails):
    from types import SimpleNamespace

    import win32gui

    from word_pdf_mcp.worker import word_window_pid

    word = SimpleNamespace(Documents=SimpleNamespace(Count=0), Caption="Original caption")

    def enumerate_windows(callback, context):
        if enumeration_fails:
            raise OSError("Window enumeration failed")

    monkeypatch.setattr(win32gui, "EnumWindows", enumerate_windows)
    with pytest.raises((OSError, RuntimeError)):
        word_window_pid(word)
    assert word.Caption == "Original caption"


def test_word_identity_rejects_instance_containing_documents():
    from types import SimpleNamespace

    from word_pdf_mcp.worker import word_window_pid

    # No Caption attribute: rejection must happen before reading or changing it.
    word = SimpleNamespace(Documents=SimpleNamespace(Count=1))
    with pytest.raises(RuntimeError, match="already contains documents"):
        word_window_pid(word)


def test_mutex_serializes_independent_processes():
    script = """
import json, time
from word_pdf_mcp.mutex import word_mutex
with word_mutex(20):
    start = time.monotonic()
    time.sleep(0.15)
    end = time.monotonic()
print(json.dumps([start, end]))
"""
    children = [
        subprocess.Popen(
            [sys.executable, "-c", script],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        for _ in range(3)
    ]
    intervals = []
    for child in children:
        output, errors = child.communicate(timeout=30)
        assert child.returncode == 0, errors.decode()
        intervals.append(json.loads(output))
    intervals.sort()
    assert all(a[1] <= b[0] for a, b in zip(intervals, intervals[1:], strict=False))


def test_abandoned_mutex_is_reported_and_released():
    from word_pdf_mcp.mutex import word_mutex

    script = """
import os
from word_pdf_mcp.mutex import word_mutex
with word_mutex(20):
    print('acquired', flush=True)
    input()
    os._exit(42)
"""
    child = subprocess.Popen(
        [sys.executable, "-c", script],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
    )
    try:
        assert child.stdout.readline().strip() == b"acquired"
        # Hold a handle open so the object persists when the owner process dies.
        import win32api
        import win32event

        from word_pdf_mcp.mutex import MUTEX_NAME

        handle = win32event.CreateMutex(None, False, MUTEX_NAME)
        try:
            child.communicate(input=b"\n", timeout=10)
            with word_mutex(5) as abandoned:
                assert abandoned is True
            with word_mutex(5) as abandoned:
                assert abandoned is False
        finally:
            win32api.CloseHandle(handle)
    finally:
        if child.poll() is None:
            child.kill()
            child.wait(timeout=10)


def test_job_close_kills_only_assigned_child():
    import win32api
    import win32con
    import win32job

    child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
    other = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
    job = None
    handle = None
    try:
        job = win32job.CreateJobObject(None, "")
        limits = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        limits["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, limits)
        handle = win32api.OpenProcess(
            win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
            False,
            child.pid,
        )
        win32job.AssignProcessToJobObject(job, handle)
        win32api.CloseHandle(job)
        job = None
        child.wait(timeout=10)
        assert other.poll() is None
    finally:
        if job is not None:
            win32api.CloseHandle(job)
        if handle is not None:
            win32api.CloseHandle(handle)
        for process in (child, other):
            if process.poll() is None:
                process.kill()
            process.wait(timeout=10)


@pytest.mark.word
def test_native_word_opt_in(tmp_path):
    source = os.environ.get("WORD_PDF_TEST_DOCUMENT")
    if not source:
        pytest.skip("Set WORD_PDF_TEST_DOCUMENT to a real local sample")
    from word_pdf_mcp.scheduler import ConversionService

    output = tmp_path / "native.pdf"
    started = time.monotonic()
    result = ConversionService().convert(source, str(output))
    assert result["status"] == "success"
    assert result["pages"] > 0
    assert result["api"] in {"ExportAsFixedFormat2", "ExportAsFixedFormat"}
    assert time.monotonic() - started < 200
    assert result["visual_fidelity_verified"] is False
