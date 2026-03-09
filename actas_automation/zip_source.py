"""ZIP parsing utilities for manuscript detection."""

from __future__ import annotations

import logging
import re
import zipfile
from shutil import copyfileobj
from pathlib import Path

from .models import SubmissionEntry
from .text_utils import clean_text

MANUSCRIPT_EXTENSIONS = {".pdf", ".doc", ".docx", ".odt"}


def read_zip_submissions(zip_path: Path, logger: logging.Logger) -> list[SubmissionEntry]:
    """Extract candidate student names from manuscript files inside the ZIP."""
    entries: list[SubmissionEntry] = []
    with zipfile.ZipFile(zip_path) as zip_file:
        for member in zip_file.namelist():
            if member.endswith("/"):
                continue
            filename = Path(member).name
            extension = Path(filename).suffix.lower()
            if extension not in MANUSCRIPT_EXTENSIONS:
                continue

            extracted_name = _extract_name_from_filename(filename)
            if not extracted_name:
                logger.warning("Skipping ZIP member without clear name: %s", member)
                continue

            entries.append(
                SubmissionEntry(
                    zip_member=member,
                    extracted_name=extracted_name,
                    extension=extension,
                )
            )
    logger.info("Manuscript files found in ZIP: %s", len(entries))
    return entries


def extract_zip_member(zip_path: Path, member: str, destination: Path) -> Path:
    """Extract one file member from ZIP into destination path."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as zip_file:
        with zip_file.open(member) as source, destination.open("wb") as target:
            copyfileobj(source, target)
    return destination


def _extract_name_from_filename(filename: str) -> str:
    stem = clean_text(Path(filename).stem)
    if not stem:
        return ""

    without_id = re.sub(r"^\d+\s*-\s*", "", stem)
    parts = [part.strip() for part in without_id.split(" - ") if part.strip()]
    candidate = parts[0] if parts else without_id
    candidate = re.sub(
        r"(?i)\b(tfm|tft|trabajo\s*fin|master|final)\b.*$",
        "",
        candidate,
    )
    return clean_text(candidate)
