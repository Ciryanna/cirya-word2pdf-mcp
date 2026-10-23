from types import SimpleNamespace

import pytest

from word_pdf_mcp.worker import export_pdf


def test_image_quality_and_markup_are_explicit():
    calls = []
    doc = SimpleNamespace(ExportAsFixedFormat2=lambda **kwargs: calls.append(kwargs))
    assert export_pdf(doc, "output.pdf", True) == "ExportAsFixedFormat2"
    assert calls[0]["OptimizeForImageQuality"] is True
    assert calls[0]["OptimizeFor"] == 0
    assert calls[0]["Range"] == 0
    assert calls[0]["Item"] == 7
    assert calls[0]["UseISO19005_1"] is False


def test_old_word_fallback_only_when_method_absent():
    calls = []
    doc = SimpleNamespace(ExportAsFixedFormat=lambda **kwargs: calls.append(kwargs))
    assert export_pdf(doc, "output.pdf", False) == "ExportAsFixedFormat"
    assert calls[0]["Item"] == 0
    assert "OptimizeForImageQuality" not in calls[0]


def test_export_failure_is_not_silently_retried():
    calls = []

    def broken(**kwargs):
        raise RuntimeError("Word export failed")

    doc = SimpleNamespace(
        ExportAsFixedFormat2=broken,
        ExportAsFixedFormat=lambda **kwargs: calls.append(kwargs),
    )
    with pytest.raises(RuntimeError, match="Word export failed"):
        export_pdf(doc, "output.pdf", False)
    assert not calls
