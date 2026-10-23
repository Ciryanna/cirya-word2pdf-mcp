from __future__ import annotations

import hashlib
import json
import os
import zipfile
from pathlib import Path

from pypdf import PdfReader

from .errors import ConversionError


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def atomic_json(path: Path, data: dict) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    with temporary.open("w", encoding="utf-8") as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def validate_paths(source: str, destination: str, overwrite: bool) -> tuple[Path, Path]:
    src = Path(source).expanduser().resolve(strict=True)
    dst = Path(destination).expanduser().resolve()
    if not src.is_file() or src.suffix.lower() not in {".doc", ".docx", ".docm"}:
        raise ConversionError("invalid_input", "Input must be a .doc, .docx or .docm file.")
    if dst.suffix.lower() != ".pdf":
        raise ConversionError("invalid_output", "Output must end in .pdf.")
    if dst.exists() and os.path.samefile(src, dst):
        raise ConversionError("same_file", "Input and output must be different files.")
    if not dst.parent.is_dir():
        raise ConversionError("invalid_output", "Create the output directory before converting.")
    if dst.exists() and not overwrite:
        raise ConversionError(
            "output_exists", "Output exists; choose another path or overwrite=True."
        )
    return src, dst


def inspect_package(path: Path) -> dict:
    """Inventory only; never parse and re-save the source through a document library."""
    if path.suffix.lower() == ".doc":
        return {"format": "doc", "warnings": ["Legacy DOC inventory is not implemented."]}
    try:
        with zipfile.ZipFile(path) as package:
            names = package.namelist()
            embeddings = [name for name in names if name.startswith("word/embeddings/")]
            return {
                "format": path.suffix.lower()[1:],
                "embedded_parts": len(embeddings),
                "warnings": (
                    ["Embedded objects found; their rendered appearance needs visual review."]
                    if embeddings
                    else []
                ),
            }
    except zipfile.BadZipFile:
        return {"warnings": ["Package inventory unavailable (encrypted or invalid OOXML file)."]}


def validate_pdf(path: Path) -> dict:
    try:
        if path.stat().st_size < 8:
            raise ValueError("PDF is empty or truncated")
        with path.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise ValueError("Missing PDF header")
            stream.seek(0)
            reader = PdfReader(stream, strict=True)
            if reader.is_encrypted:
                raise ValueError("Encrypted PDF is not supported by the basic validator")
            pages = len(reader.pages)
            if not pages:
                raise ValueError("PDF has no pages")
            sizes = [
                [float(page.mediabox.width), float(page.mediabox.height)] for page in reader.pages
            ]
        return {"pages": pages, "page_sizes_pt": sizes, "bytes": path.stat().st_size}
    except Exception as exc:
        raise ConversionError("invalid_pdf", str(exc)) from exc


def publish_pdf(staged: Path, destination: Path, overwrite: bool) -> None:
    if overwrite:
        os.replace(staged, destination)
    elif os.name == "nt":
        # Windows rename fails if the destination already exists (no check-then-replace race).
        os.rename(staged, destination)
    else:
        # Used by portable tests; staging is on the same filesystem as destination.
        os.link(staged, destination)
        staged.unlink()
