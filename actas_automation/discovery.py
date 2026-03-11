"""Workspace discovery helpers."""

from __future__ import annotations

import logging
from pathlib import Path

from pypdf import PdfReader

from .models import SourceFiles
from .text_utils import normalize_text


def discover_source_files(workspace_dir: Path, logger: logging.Logger) -> SourceFiles:
    """Find the ZIP, Excel, template, and PDF files needed by the pipeline."""
    zip_file = _discover_zip_file(workspace_dir)
    excel_file = _discover_excel_file(workspace_dir)
    acta_pdf_template = _discover_acta_pdf_template(workspace_dir, logger)
    pdf_files = [
        path
        for path in _discover_pdf_files(workspace_dir)
        if path.resolve() != acta_pdf_template.resolve()
    ]

    logger.info("ZIP detected: %s", zip_file.name)
    logger.info("Excel detected: %s", excel_file.name)
    logger.info("Acta PDF template detected: %s", acta_pdf_template.name)
    logger.info("PDF files detected: %s", len(pdf_files))

    return SourceFiles(
        excel_file=excel_file,
        zip_file=zip_file,
        acta_pdf_template=acta_pdf_template,
        pdf_files=pdf_files,
    )


def _discover_zip_file(workspace_dir: Path) -> Path:
    zip_candidates = [
        path
        for path in workspace_dir.glob("*.zip")
        if not path.name.startswith("~$") and not path.name.startswith(".")
    ]
    if not zip_candidates:
        raise FileNotFoundError("No ZIP file found in workspace root.")
    zip_candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return zip_candidates[0]


def _discover_excel_file(workspace_dir: Path) -> Path:
    excel_candidates = [
        path
        for path in workspace_dir.glob("*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not excel_candidates:
        raise FileNotFoundError("No .xlsx file found in workspace.")
    excel_candidates.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    return excel_candidates[0]


def _discover_acta_pdf_template(
    workspace_dir: Path,
    logger: logging.Logger,
) -> Path:
    pdf_candidates = [
        path
        for path in workspace_dir.glob("*.pdf")
        if not path.name.startswith("~$") and not path.name.startswith("._")
    ]
    if not pdf_candidates:
        raise FileNotFoundError(
            "No acta PDF template found in workspace root (*.pdf)."
        )

    scored: list[tuple[int, Path]] = []
    for path in pdf_candidates:
        try:
            reader = PdfReader(str(path), strict=False)
            fields = reader.get_fields() or {}
            text = "\n".join(
                (page.extract_text() or "") for page in reader.pages[:1]
            )
            normalized = normalize_text(text)
            score = 0
            if fields:
                score += 10
            for marker in ("acta", "tribunal", "datos del estudiante"):
                if marker in normalized:
                    score += 2
            if "apellido" in normalize_text(path.stem):
                score += 1
            scored.append((score, path))
        except Exception as exc:  # pragma: no cover - best effort classification
            logger.warning("Could not inspect PDF template %s: %s", path.name, exc)

    if not scored:
        raise FileNotFoundError("No readable acta PDF template found in workspace root.")

    scored.sort(key=lambda item: item[0], reverse=True)
    best_score, best_path = scored[0]
    if best_score < 10:
        raise FileNotFoundError(
            "No fillable acta PDF template found in workspace root."
        )
    return best_path


def _discover_pdf_files(workspace_dir: Path) -> list[Path]:
    pdf_files: list[Path] = []
    for path in workspace_dir.rglob("*.pdf"):
        if _should_skip_pdf(path, workspace_dir):
            continue
        pdf_files.append(path)
    return sorted(pdf_files)


def _should_skip_pdf(path: Path, workspace_dir: Path) -> bool:
    try:
        relative_parts = path.resolve().relative_to(workspace_dir.resolve()).parts
    except ValueError:
        return True

    if any(part.startswith(".") for part in relative_parts):
        return True
    if any(part.startswith("output") for part in relative_parts):
        return True

    name = path.name
    if name.startswith("~$") or name.startswith("._"):
        return True
    return False
