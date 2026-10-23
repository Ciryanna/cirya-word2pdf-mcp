from __future__ import annotations

import sys
import time
from contextlib import contextmanager

from .errors import ConversionError

MUTEX_NAME = r"Local\Elwina.WordPdf.Worker.v1"


@contextmanager
def word_mutex(timeout_seconds: float):
    """The calling OS thread must both acquire and release this mutex.

    Local\\ coordinates MCP instances in the same Windows logon session.
    WAIT_ABANDONED is ownership *plus* a recovery signal, not an ordinary failure.
    """
    if sys.platform != "win32":
        raise ConversionError("windows_required", "Native Word export requires Windows.")
    import win32api
    import win32event

    handle = win32event.CreateMutex(None, False, MUTEX_NAME)
    acquired = False
    try:
        deadline = time.monotonic() + timeout_seconds
        while True:
            remaining = max(0, deadline - time.monotonic())
            result = win32event.WaitForSingleObject(handle, min(250, int(remaining * 1000)))
            if result in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
                acquired = True
                yield result == win32event.WAIT_ABANDONED
                break
            if result != win32event.WAIT_TIMEOUT:
                raise ConversionError("mutex_failed", f"Unexpected mutex result: {result}")
            if time.monotonic() >= deadline:
                raise ConversionError(
                    "queue_timeout", "Timed out waiting for the Word export lock."
                )
    finally:
        try:
            if acquired:
                win32event.ReleaseMutex(handle)
        finally:
            win32api.CloseHandle(handle)
