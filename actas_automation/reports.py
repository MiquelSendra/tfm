"""PDF parsing for director reports and optional PDF template fallback."""

from __future__ import annotations

import logging
import re
from pathlib import Path

from pypdf import PdfReader

from .models import ActaContext, DirectorReport
from .text_utils import clean_text, normalize_text


def parse_pdf_files(
    pdf_files: list[Path], logger: logging.Logger
) -> tuple[list[DirectorReport], list[Path]]:
    """Classify PDFs into director reports and non-report candidates."""
    reports: list[DirectorReport] = []
    non_report_pdfs: list[Path] = []

    for path in pdf_files:
        text = extract_pdf_text(path, logger)
        if not text:
            non_report_pdfs.append(path)
            continue

        if is_director_report(text):
            report = _parse_director_report(path, text)
            if report.extracted_name:
                reports.append(report)
            else:
                logger.warning("Could not extract student name from report: %s", path.name)
        else:
            non_report_pdfs.append(path)

    logger.info("Director reports detected: %s", len(reports))
    return reports, non_report_pdfs


def extract_pdf_text(path: Path, logger: logging.Logger) -> str:
    """Read all extractable text from a PDF file."""
    try:
        reader = PdfReader(str(path), strict=False)
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    except Exception as exc:  # pragma: no cover - depends on third-party parser
        logger.warning("Could not read PDF text from %s: %s", path.name, exc)
        return ""


def is_director_report(text: str) -> bool:
    """Heuristic detector for director report PDFs."""
    normalized = normalize_text(text)
    return (
        "informe del director" in normalized
        and "datos del la estudiante" in normalized
    ) or (
        "informe del director" in normalized and "trabajo fin de titulo" in normalized
    )


def _parse_director_report(path: Path, text: str) -> DirectorReport:
    student_section = _extract_section(
        text,
        start_markers=("Datos del/la estudiante", "Datos del estudiante"),
        end_markers=("Datos del/la Director/a", "Datos del Director/a"),
    )
    title_section = _extract_section(
        text,
        start_markers=(
            "Título del Trabajo Fin de Título (TFT)",
            "Titulo del Trabajo Fin de Titulo (TFT)",
        ),
        end_markers=("Datos del/la Director/a", "Datos del Director/a"),
    )

    surnames = _extract_line_value(student_section, "Apellidos")
    given_name = _extract_line_value(student_section, "Nombre")
    full_name = ""
    if surnames and given_name:
        full_name = f"{surnames}, {given_name}"
    elif given_name:
        full_name = given_name
    elif surnames:
        full_name = surnames

    thesis_title = _clean_multiline_text(title_section)
    return DirectorReport(
        path=path,
        extracted_name=clean_text(full_name),
        thesis_title=thesis_title,
        extracted_text=text,
    )


def _extract_section(text: str, start_markers: tuple[str, ...], end_markers: tuple[str, ...]) -> str:
    start_index = -1
    for marker in start_markers:
        pos = text.lower().find(marker.lower())
        if pos != -1:
            start_index = pos + len(marker)
            break
    if start_index == -1:
        return text

    section = text[start_index:]
    end_index = len(section)
    for marker in end_markers:
        pos = section.lower().find(marker.lower())
        if pos != -1:
            end_index = min(end_index, pos)
    return section[:end_index]


def _extract_line_value(section_text: str, key: str) -> str:
    pattern = re.compile(rf"{re.escape(key)}\s*:\s*([^\n\r]+)", re.IGNORECASE)
    match = pattern.search(section_text)
    if not match:
        return ""
    return clean_text(match.group(1))


def _clean_multiline_text(value: str) -> str:
    lines = [clean_text(line) for line in value.splitlines()]
    lines = [line for line in lines if line]
    return " ".join(lines).strip()


def try_fill_pdf_template(
    template_pdf: Path,
    output_pdf: Path,
    context: ActaContext,
    logger: logging.Logger,
) -> bool:
    """Fill a PDF form template if AcroForm fields exist; return success."""
    try:
        from pypdf import PdfReader, PdfWriter
        from pypdf.generic import NameObject, TextStringObject
    except Exception:  # pragma: no cover - guarded import
        return False

    reader = PdfReader(str(template_pdf), strict=False)
    fields = reader.get_fields() or {}
    if not fields:
        logger.warning("PDF template has no form fields: %s", template_pdf.name)
        return False

    field_values = _map_pdf_fields(fields.keys(), context)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    for page in writer.pages:
        annotations = page.get("/Annots") or []
        for annotation in annotations:
            widget = annotation.get_object()
            field_name = widget.get("/T")
            if field_name not in field_values:
                continue
            value = field_values[field_name]
            widget[NameObject("/V")] = TextStringObject(value)
            widget[NameObject("/DV")] = TextStringObject(value)

    writer.set_need_appearances_writer(True)

    with output_pdf.open("wb") as handle:
        writer.write(handle)
    return True


def _map_pdf_fields(field_names: list[str], context: ActaContext) -> dict[str, str]:
    value_map: dict[str, str] = {}
    for name in field_names:
        normalized = normalize_text(name)
        if any(
            token in normalized
            for token in ("presidente", "secretario", "apto", "observaciones")
        ):
            continue
        if normalized.startswith("date") or normalized.startswith("text4"):
            continue

        if "titulacion" in normalized:
            value_map[name] = context.titulacion
        elif "por ejemplo" in normalized and (
            "master" in normalized or "bioinformatica" in normalized
        ):
            value_map[name] = context.titulacion
        elif "curso" in normalized or "edicion" in normalized:
            value_map[name] = context.edicion
        elif "por ejemplo" in normalized and "abril" in normalized:
            value_map[name] = context.edicion
        elif "dni" in normalized or "pasaporte" in normalized:
            if "nombre apellidos y dniniepasaporte" in normalized:
                value_map[name] = context.director
            else:
                value_map[name] = context.dni
        elif "nombre" in normalized and "director" not in normalized:
            if "presidente" in normalized or "secretario" in normalized:
                continue
            value_map[name] = context.student_name
        elif "titulo" in normalized and "trabajo" in normalized:
            value_map[name] = context.thesis_title
        elif "titulo del tft" in normalized:
            value_map[name] = context.thesis_title
        elif "director" in normalized:
            value_map[name] = context.director
        elif "nombre apellidos y dniniepasaporte" in normalized:
            value_map[name] = context.director
    return value_map
