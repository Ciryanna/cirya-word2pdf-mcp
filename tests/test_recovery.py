import json
from types import SimpleNamespace

import pytest

from word_pdf_mcp import processes, scheduler
from word_pdf_mcp.errors import ConversionError


def test_unknown_word_ownership_blocks_following_tasks(tmp_path):
    (tmp_path / "state.json").write_text(json.dumps({"phase": "starting_word"}))
    with pytest.raises(ConversionError, match="recovery_required"):
        scheduler.cleanup_worker({"task_dir": str(tmp_path)})


def test_before_word_crash_needs_no_word_kill(tmp_path, monkeypatch):
    (tmp_path / "state.json").write_text(json.dumps({"phase": "before_word"}))
    calls = []
    monkeypatch.setattr(scheduler, "terminate_owned_word", lambda record: calls.append(record))
    scheduler.cleanup_worker({"task_dir": str(tmp_path)})
    assert not calls


def test_identified_word_cleanup_uses_recorded_identity(tmp_path, monkeypatch):
    record = {"pid": 1234, "created": 123.5, "name": "WINWORD.EXE"}
    (tmp_path / "state.json").write_text(json.dumps({"phase": "job_assigned", "word": record}))
    calls = []
    monkeypatch.setattr(scheduler, "terminate_owned_word", lambda value: calls.append(value))
    scheduler.cleanup_worker({"task_dir": str(tmp_path)})
    assert calls == [record]


def test_pid_reuse_must_not_kill_unrelated_process(monkeypatch):
    calls = []
    new_process = SimpleNamespace(
        create_time=lambda: 900,
        name=lambda: "WINWORD.EXE",
        kill=lambda: calls.append("kill"),
    )
    monkeypatch.setattr(processes.psutil, "Process", lambda pid: new_process)
    processes.terminate_owned_word({"pid": 123, "created": 100, "name": "WINWORD.EXE"})
    assert not calls
