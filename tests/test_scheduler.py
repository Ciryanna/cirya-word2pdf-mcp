import json
import subprocess
import threading
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from types import SimpleNamespace

import pytest
from pypdf import PdfWriter

from word_pdf_mcp import scheduler
from word_pdf_mcp.errors import ConversionError


@pytest.fixture
def harness(tmp_path, monkeypatch):
    source = tmp_path / "source.doc"
    source.write_bytes(b"opaque original bytes")
    root = tmp_path / "state"
    root.mkdir()
    service = scheduler.ConversionService(root)
    events = []

    @contextmanager
    def mutex(timeout):
        events.append("acquire")
        try:
            yield False
        finally:
            events.append("release")

    monkeypatch.setattr(scheduler.sys, "platform", "win32")
    monkeypatch.setattr(scheduler, "word_mutex", mutex)
    monkeypatch.setattr(scheduler, "identity", lambda pid: {"pid": pid})
    monkeypatch.setattr(scheduler, "cleanup_worker", lambda pending: events.append("cleanup"))
    return SimpleNamespace(
        source=source,
        output=tmp_path / "out.pdf",
        root=root,
        service=service,
        events=events,
    )


def fake_worker(monkeypatch, *, pages=1, reported_pages=1, timeout=False):
    class Process:
        pid = 123456
        returncode = None

        def __init__(self, args, **kwargs):
            request = json.loads(open(args[-3], encoding="utf-8").read())
            writer = PdfWriter()
            for _ in range(pages):
                writer.add_blank_page(595, 842)
            with open(request["output_path"], "wb") as output:
                writer.write(output)
            with open(args[-1], "w", encoding="utf-8") as result:
                json.dump({"status": "success", "word_pages": reported_pages}, result)

        def wait(self, timeout=None):
            if timeout is not None and globals_timeout and self.returncode is None:
                raise subprocess.TimeoutExpired("fake", timeout)
            self.returncode = 0 if self.returncode is None else self.returncode
            return self.returncode

        def poll(self):
            return self.returncode

        def kill(self):
            self.returncode = -9

    globals_timeout = timeout
    monkeypatch.setattr(scheduler.subprocess, "Popen", Process)


def test_cleanup_finishes_before_lock_release_and_source_unchanged(harness, monkeypatch):
    fake_worker(monkeypatch)
    result = harness.service.convert(str(harness.source), str(harness.output))
    assert result["pages"] == 1
    assert result["visual_fidelity_verified"] is False
    assert harness.source.read_bytes() == b"opaque original bytes"
    assert harness.events == ["acquire", "cleanup", "release"]
    assert not (harness.root / "pending.json").exists()
    assert list(harness.root.iterdir()) == []


def test_timeout_cleans_before_release_and_preserves_existing_output(harness, monkeypatch):
    fake_worker(monkeypatch, timeout=True)
    harness.output.write_bytes(b"previous valid user output")
    with pytest.raises(ConversionError, match="conversion_timeout"):
        harness.service.convert(str(harness.source), str(harness.output), overwrite=True)
    assert harness.output.read_bytes() == b"previous valid user output"
    assert harness.events == ["acquire", "cleanup", "release"]


def test_pdf_page_mismatch_does_not_publish(harness, monkeypatch):
    fake_worker(monkeypatch, pages=2, reported_pages=1)
    with pytest.raises(ConversionError, match="page_count_mismatch"):
        harness.service.convert(str(harness.source), str(harness.output))
    assert not harness.output.exists()


def test_failed_cleanup_leaves_journal_and_blocks_next_launch(harness, monkeypatch):
    fake_worker(monkeypatch)

    def cannot_cleanup(pending):
        raise RuntimeError("Cannot identify Word")

    monkeypatch.setattr(scheduler, "cleanup_worker", cannot_cleanup)
    with pytest.raises(ConversionError, match="recovery_required"):
        harness.service.convert(str(harness.source), str(harness.output))
    assert (harness.root / "pending.json").exists()
    assert not harness.output.exists()
    # Replacing Popen ensures recovery failure cannot sneak into another Word launch.
    monkeypatch.setattr(scheduler.subprocess, "Popen", lambda *a, **kw: pytest.fail("new worker"))
    with pytest.raises(RuntimeError, match="Cannot identify Word"):
        harness.service.convert(str(harness.source), str(harness.output))


def test_concurrent_calls_cannot_enter_word_until_previous_finishes(harness, monkeypatch):
    first_entered = threading.Event()
    second_entered = threading.Event()
    release_first = threading.Event()
    second_requested = threading.Event()
    calls = []

    def run(src, dst, *args, **kwargs):
        calls.append(dst.name)
        if len(calls) == 1:
            first_entered.set()
            assert release_first.wait(5)
        else:
            second_entered.set()
        return {"status": "success"}

    def second():
        second_requested.set()
        return harness.service.convert(str(harness.source), str(harness.output))

    monkeypatch.setattr(harness.service, "_run", run)
    with ThreadPoolExecutor(max_workers=2) as pool:
        first = pool.submit(harness.service.convert, str(harness.source), str(harness.output))
        try:
            assert first_entered.wait(5)
            following = pool.submit(second)
            assert second_requested.wait(5)
            assert not second_entered.wait(0.1)
        finally:
            release_first.set()
        assert first.result(timeout=5)["status"] == "success"
        assert following.result(timeout=5)["status"] == "success"
    assert second_entered.is_set()


def test_failed_worker_launch_does_not_leave_false_recovery_block(harness, monkeypatch):
    def fail(*args, **kwargs):
        raise OSError("Cannot start Python")

    monkeypatch.setattr(scheduler.subprocess, "Popen", fail)
    with pytest.raises(ConversionError, match="worker_start_failed"):
        harness.service.convert(str(harness.source), str(harness.output))
    assert not (harness.root / "pending.json").exists()
    assert list(harness.root.iterdir()) == []
