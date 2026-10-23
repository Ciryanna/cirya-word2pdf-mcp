from __future__ import annotations

import time

import psutil


def identity(pid: int) -> dict:
    process = psutil.Process(pid)
    return {"pid": pid, "created": process.create_time(), "name": process.name()}


def matching_process(record: dict | None) -> psutil.Process | None:
    if not record:
        return None
    try:
        process = psutil.Process(int(record["pid"]))
        if abs(process.create_time() - float(record["created"])) > 0.01:
            return None  # PID was reused; never terminate that new process.
        if process.name().lower() != str(record["name"]).lower():
            return None
        if not process.is_running() or process.status() == psutil.STATUS_ZOMBIE:
            return None
        return process
    except psutil.NoSuchProcess:
        return None


def wait_gone(record: dict | None, timeout: float = 5) -> bool:
    deadline = time.monotonic() + timeout
    while matching_process(record) is not None:
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.05)
    return True


def terminate_owned_word(record: dict) -> None:
    """Only act on a PID+creation-time identity positively verified by our worker."""
    if record.get("name", "").lower() != "winword.exe":
        raise ValueError("Refusing to terminate a process that is not WINWORD.EXE")
    process = matching_process(record)
    if process is not None:
        process.kill()
    if not wait_gone(record):
        raise RuntimeError("Owned Word process did not exit")
