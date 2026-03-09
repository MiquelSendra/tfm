"""Excel parsing for metadata and student records."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Iterable

import pandas as pd

from .models import ProgramMetadata, StudentRecord
from .text_utils import build_name_aliases, clean_text, normalize_text


def load_students_and_metadata(
    excel_path: str | Path, logger: logging.Logger
) -> tuple[ProgramMetadata, list[StudentRecord]]:
    """Load metadata and student rows from the first sheet of the Excel file."""
    raw_df = pd.read_excel(excel_path, header=None, dtype=str)
    metadata = _extract_metadata(raw_df)
    header_row_index = _detect_header_row(raw_df)
    if header_row_index is None:
        raise ValueError("Could not detect table header row in Excel.")

    headers = _build_headers(raw_df.iloc[header_row_index].tolist())
    table_df = raw_df.iloc[header_row_index + 1 :].copy()
    table_df.columns = headers
    table_df = table_df.dropna(how="all")

    students = _build_student_records(table_df, header_row_index + 2, logger)
    logger.info("Excel table header row detected at row %s", header_row_index + 1)
    logger.info("Student rows parsed: %s", len(students))
    logger.info(
        "Metadata extracted: titulo='%s' | edicion='%s'",
        metadata.title,
        metadata.edition,
    )
    return metadata, students


def _extract_metadata(raw_df: pd.DataFrame) -> ProgramMetadata:
    title = ""
    edition = ""
    title_pattern = re.compile(r"^\s*t[ií]tulo\s*:\s*(.+)$", re.IGNORECASE)
    edition_pattern = re.compile(r"^\s*edici[oó]n\s*:\s*(.+)$", re.IGNORECASE)

    for value in raw_df.values.flatten():
        text = clean_text(value)
        if not text:
            continue
        title_match = title_pattern.match(text)
        if title_match and not title:
            title = title_match.group(1).strip()
            continue
        edition_match = edition_pattern.match(text)
        if edition_match and not edition:
            edition = edition_match.group(1).strip()

    return ProgramMetadata(title=title, edition=edition)


def _detect_header_row(raw_df: pd.DataFrame) -> int | None:
    required_tokens = ("nombre", "dni", "director", "mail")
    best_idx: int | None = None
    best_score = -1

    for idx, row in raw_df.iterrows():
        cells = [normalize_text(clean_text(value)) for value in row.tolist()]
        cells = [cell for cell in cells if cell]
        if not cells:
            continue

        score = 0
        for token in required_tokens:
            if any(token in cell for cell in cells):
                score += 1
        if any("titulo" in cell and ("tfm" in cell or "tfg" in cell) for cell in cells):
            score += 1
        if any("tema" in cell and "tft" in cell for cell in cells):
            score += 1

        if score > best_score:
            best_score = score
            best_idx = idx

    if best_score < 4:
        return None
    return best_idx


def _build_headers(raw_headers: Iterable[object]) -> list[str]:
    headers: list[str] = []
    used: dict[str, int] = {}

    for position, value in enumerate(raw_headers):
        text = normalize_text(clean_text(value))
        if not text:
            text = f"col_{position + 1}"
        if text in used:
            used[text] += 1
            text = f"{text}_{used[text]}"
        else:
            used[text] = 1
        headers.append(text)

    return headers


def _build_student_records(
    table_df: pd.DataFrame, data_start_row: int, logger: logging.Logger
) -> list[StudentRecord]:
    name_col = _find_col(table_df.columns, include=("nombre", "completo"))
    dni_col = _find_col(table_df.columns, include=("dni",))
    email_col = _find_col(table_df.columns, include=("mail",))
    director_col = _find_col(
        table_df.columns,
        include=("director",),
        exclude=("correo", "mail"),
    )
    title_col = _find_col(table_df.columns, include=("titulo",), optional=True)
    topic_col = _find_col(table_df.columns, include=("tema", "tft"), optional=True)

    if not name_col or not dni_col:
        raise ValueError("Excel columns for student name and DNI are required.")

    students: list[StudentRecord] = []
    for idx, row in table_df.iterrows():
        full_name = clean_text(row.get(name_col, ""))
        dni = clean_text(row.get(dni_col, ""))
        if not full_name:
            continue

        record = StudentRecord(
            full_name=full_name,
            dni=dni,
            email=clean_text(row.get(email_col, "")),
            director=clean_text(row.get(director_col, "")),
            thesis_title=clean_text(row.get(title_col, "")),
            thesis_topic=clean_text(row.get(topic_col, "")),
            source_row=data_start_row + (idx - table_df.index[0]),
            aliases=build_name_aliases(full_name),
        )
        students.append(record)

    if not students:
        logger.warning("No student rows were parsed from the Excel table.")
    return students


def _find_col(
    columns: Iterable[str],
    include: tuple[str, ...],
    exclude: tuple[str, ...] = (),
    optional: bool = False,
) -> str:
    for col in columns:
        if all(token in col for token in include) and not any(token in col for token in exclude):
            return col
    if optional:
        return ""
    raise ValueError(
        f"Could not locate required column with include={include} exclude={exclude}."
    )
