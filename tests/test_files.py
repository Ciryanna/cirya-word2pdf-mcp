import zipfile

import pytest
from pypdf import PdfWriter

from word_pdf_mcp.errors import ConversionError
from word_pdf_mcp.files import inspect_package, publish_pdf, sha256, validate_paths, validate_pdf


def write_pdf(path, pages=1):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=595, height=842)
    with path.open("wb") as stream:
        writer.write(stream)


def test_pdf_validation_detects_truncation_and_reads_geometry(tmp_path):
    pdf = tmp_path / "output.pdf"
    pdf.write_bytes(b"%PDF-1.7\ntruncated")
    with pytest.raises(ConversionError, match="invalid_pdf"):
        validate_pdf(pdf)
    write_pdf(pdf, pages=2)
    result = validate_pdf(pdf)
    assert result["pages"] == 2
    assert result["page_sizes_pt"] == [[595, 842], [595, 842]]


def test_existing_output_never_overwritten_even_if_it_appears_after_validation(tmp_path):
    src = tmp_path / "input.docx"
    src.write_bytes(b"source")
    output = tmp_path / "output.pdf"
    validate_paths(str(src), str(output), False)
    output.write_bytes(b"concurrent user's output")
    staged = tmp_path / "stage.pdf"
    write_pdf(staged)
    with pytest.raises(FileExistsError):
        publish_pdf(staged, output, False)
    assert output.read_bytes() == b"concurrent user's output"


def test_explicit_overwrite_replaces_atomically(tmp_path):
    output, staged = tmp_path / "output.pdf", tmp_path / "stage.pdf"
    output.write_bytes(b"old")
    write_pdf(staged)
    expected = sha256(staged)
    publish_pdf(staged, output, True)
    assert sha256(output) == expected
    assert not staged.exists()


def test_inventory_does_not_modify_embedded_parts(tmp_path):
    src = tmp_path / "input.docx"
    with zipfile.ZipFile(src, "w") as package:
        package.writestr("word/document.xml", "<document/>")
        package.writestr("word/embeddings/oleObject1.bin", b"opaque object data")
    original = sha256(src)
    assert inspect_package(src)["embedded_parts"] == 1
    assert sha256(src) == original
