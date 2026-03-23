"""Helpers for reports and presentation files copied into student folders."""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
import logging
import re
from pathlib import Path
from shutil import copy2
import zipfile

from pypdf import PdfReader

from .text_utils import clean_text
from .zip_source import extract_zip_member


@dataclass(frozen=True)
class SupportingDocument:
    """One document that can be copied into a student folder."""

    source_file: str
    candidate_name: str
    email: str = ""
    path: Path | None = None
    zip_path: Path | None = None
    zip_member: str = ""


def read_slide_documents(slides_dir: Path, logger: logging.Logger) -> list[SupportingDocument]:
    """Read slide PDFs from a ZIP bundle or directly from the folder."""
    if not slides_dir.exists():
        logger.info("Slides directory not found: %s", slides_dir)
        return []

    documents: list[SupportingDocument] = []

    for path in sorted(slides_dir.rglob("*.pdf")):
        if _should_skip_path(path):
            continue
        first_page_text = _extract_first_page_text_from_path(path, logger)
        documents.append(
            SupportingDocument(
                source_file=str(path),
                candidate_name=extract_student_name_from_cover(first_page_text),
                email=_extract_email_hint(path.name),
                path=path,
            )
        )

    for zip_path in sorted(slides_dir.rglob("*.zip")):
        if _should_skip_path(zip_path):
            continue
        documents.extend(_read_slide_documents_from_zip(zip_path, logger))

    logger.info("Slide documents detected: %s", len(documents))
    return documents


def copy_supporting_document(
    document: SupportingDocument,
    destination: Path,
) -> Path:
    """Copy or extract a supporting document into the destination path."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    if document.path:
        copy2(document.path, destination)
        return destination
    if document.zip_path and document.zip_member:
        extract_zip_member(document.zip_path, document.zip_member, destination)
        return destination
    raise ValueError(f"unsupported_document_source:{document.source_file}")


def extract_student_name_from_cover(text: str) -> str:
    """Extract student name from a standard cover page."""
    cleaned_text = clean_text(text)
    if not cleaned_text:
        return ""

    patterns = (
        r"\bAlumno/a\s*:\s*(.+?)(?=\bDirector(?:/a)?\b|\bTutora?\b|$)",
        r"\bAlumn[oa]\s*:\s*(.+?)(?=\bDirector(?:/a)?\b|\bTutora?\b|$)",
        r"\bPresentado por\s*:\s*(.+?)(?=\bDirector(?:/a)?\b|\bTutora?\b|$)",
        r"\bAutor(?:a)?\s*:\s*(.+?)(?=\bDirector(?:/a)?\b|\bTutora?\b|$)",
    )

    for pattern in patterns:
        match = re.search(pattern, cleaned_text, flags=re.IGNORECASE | re.DOTALL)
        if not match:
            continue
        candidate = _cleanup_extracted_name(match.group(1))
        if candidate:
            return candidate

    lines = [_cleanup_cover_line(line) for line in cleaned_text.splitlines()]
    lines = [line for line in lines if line]
    for index, line in enumerate(lines):
        if not re.search(r"\bDirector(?:a|/a)?\b|\bTutor(?:a|/a)?\b", line, re.IGNORECASE):
            continue
        if index == 0:
            continue
        candidate = _cleanup_extracted_name(lines[index - 1])
        if _looks_like_person_name(candidate):
            return candidate

    return ""


def _read_slide_documents_from_zip(
    zip_path: Path,
    logger: logging.Logger,
) -> list[SupportingDocument]:
    documents: list[SupportingDocument] = []
    with zipfile.ZipFile(zip_path) as zip_file:
        for member in zip_file.namelist():
            if member.endswith("/") or Path(member).suffix.lower() != ".pdf":
                continue
            first_page_text = _extract_first_page_text_from_zip_member(
                zip_file,
                member,
                logger,
            )
            documents.append(
                SupportingDocument(
                    source_file=f"{zip_path.name}:{member}",
                    candidate_name=extract_student_name_from_cover(first_page_text),
                    email=_extract_email_hint(Path(member).name),
                    zip_path=zip_path,
                    zip_member=member,
                )
            )
    return documents


def _extract_first_page_text_from_path(path: Path, logger: logging.Logger) -> str:
    try:
        reader = PdfReader(str(path), strict=False)
        if not reader.pages:
            return ""
        return reader.pages[0].extract_text() or ""
    except Exception as exc:  # pragma: no cover - third-party parsing variability
        logger.warning("Could not read first page from %s: %s", path.name, exc)
        return ""


def _extract_first_page_text_from_zip_member(
    zip_file: zipfile.ZipFile,
    member: str,
    logger: logging.Logger,
) -> str:
    try:
        data = zip_file.read(member)
        reader = PdfReader(BytesIO(data), strict=False)
        if not reader.pages:
            return ""
        return reader.pages[0].extract_text() or ""
    except Exception as exc:  # pragma: no cover - third-party parsing variability
        logger.warning("Could not read first page from ZIP member %s: %s", member, exc)
        return ""


def _cleanup_extracted_name(value: str) -> str:
    candidate = clean_text(value)
    candidate = candidate.replace("|", " ")
    candidate = re.sub(r"\b\d+\s*$", "", candidate).strip()
    candidate = re.sub(r"\s+", " ", candidate)
    candidate = re.sub(
        r"\s*(Directora?|Director/a|Tutora?|Tutor/a)\s*:.*$",
        "",
        candidate,
        flags=re.IGNORECASE | re.DOTALL,
    ).strip()
    return candidate


def _extract_email_hint(filename: str) -> str:
    match = re.search(r"([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})", filename, re.IGNORECASE)
    if not match:
        return ""
    return clean_text(match.group(1)).lower()


def _cleanup_cover_line(value: str) -> str:
    return re.sub(r"\s+", " ", clean_text(value).replace("|", " ")).strip()


def _looks_like_person_name(value: str) -> bool:
    words = [part for part in re.split(r"\s+", clean_text(value)) if part]
    if len(words) < 2 or len(words) > 8:
        return False
    lowered = {word.lower() for word in words}
    forbidden = {
        "trabajo",
        "master",
        "máster",
        "bioinformatica",
        "bioinformática",
        "directora",
        "director",
        "convocatoria",
    }
    return not lowered.intersection(forbidden)


def _should_skip_path(path: Path) -> bool:
    name = path.name
    return name.startswith("~$") or name.startswith("._") or name.startswith(".")
