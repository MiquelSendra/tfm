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


def extract_pdf_page_text(path: Path, page_index: int, logger: logging.Logger) -> str:
    """Read extractable text from one PDF page."""
    try:
        reader = PdfReader(str(path), strict=False)
        if page_index < 0 or page_index >= len(reader.pages):
            return ""
        return reader.pages[page_index].extract_text() or ""
    except Exception as exc:  # pragma: no cover - depends on third-party parser
        logger.warning("Could not read PDF page %s from %s: %s", page_index + 1, path.name, exc)
        return ""


def extract_manuscript_title_from_pdf(path: Path, logger: logging.Logger) -> str:
    """Extract thesis title from the first manuscript page."""
    first_page_text = extract_pdf_page_text(path, 0, logger)
    if not first_page_text:
        return ""
    return _extract_manuscript_title_from_first_page(first_page_text)


def is_director_report(text: str) -> bool:
    """Heuristic detector for director report PDFs."""
    normalized = normalize_text(text)
    header_markers = (
        "informe del director" in normalized
        or "informe del director a" in normalized
    )
    structure_markers = (
        "datos del la estudiante" in normalized
        and "datos del la director a" in normalized
        and "titulo del trabajo fin de titulo" in normalized
    )
    evaluation_markers = (
        "presentacion del trabajo del la estudiante" in normalized
        or "atendiendo a los criterios de evaluacion" in normalized
        or "favorable" in normalized
    )
    return (header_markers and ("trabajo fin de titulo" in normalized or structure_markers)) or (
        structure_markers and evaluation_markers
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

    surnames = _extract_field_value(
        student_section,
        "Apellidos",
        stop_labels=(
            "Nombre",
            "Título del Trabajo Fin de Título (TFT)",
            "Titulo del Trabajo Fin de Titulo (TFT)",
            "Datos del/la Director/a",
            "Datos del Director/a",
            "Observaciones",
        ),
    )
    given_name = _extract_field_value(
        student_section,
        "Nombre",
        stop_labels=(
            "Título del Trabajo Fin de Título (TFT)",
            "Titulo del Trabajo Fin de Titulo (TFT)",
            "Datos del/la Director/a",
            "Datos del Director/a",
            "Observaciones",
        ),
    )
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


def _extract_field_value(
    section_text: str,
    label: str,
    stop_labels: tuple[str, ...] = (),
) -> str:
    escaped_stops = [re.escape(stop_label) for stop_label in stop_labels if stop_label]
    extra_stops = [
        r"Apellidos\s*:",
        r"Nombre\s*:",
    ]
    lookahead_options = escaped_stops + extra_stops + [r"$"]
    pattern = re.compile(
        rf"{re.escape(label)}\s*:\s*(.+?)(?=(?:{'|'.join(lookahead_options)}))",
        re.IGNORECASE | re.DOTALL,
    )
    match = pattern.search(section_text)
    if not match:
        return _extract_line_value(section_text, label)
    return _clean_inline_field(match.group(1))


def _clean_multiline_text(value: str) -> str:
    lines = [clean_text(line) for line in value.splitlines()]
    lines = [line for line in lines if line]
    return " ".join(lines).strip()


def _clean_inline_field(value: str) -> str:
    cleaned = clean_text(value)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = re.sub(r"\s*[|]+\s*", " ", cleaned)
    return cleaned.strip(" :;-")


def _extract_manuscript_title_from_first_page(text: str) -> str:
    metadata_markers = (
        "titulacion",
        "alumno",
        "alumna",
        "convocatoria",
        "curso academico",
        "dni",
        "d n i",
        "director",
        "orientacion",
        "creditos",
        "ciudad mes y ano",
    )
    title_lines: list[str] = []

    for raw_line in text.splitlines():
        line = _collapse_whitespace(clean_text(raw_line))
        if not line:
            continue

        normalized = normalize_text(line)
        if any(normalized.startswith(marker) for marker in metadata_markers):
            break
        title_lines.append(line)

    while title_lines and _looks_like_cover_preamble(title_lines[0]):
        title_lines.pop(0)

    return _collapse_whitespace(" ".join(title_lines))


def _looks_like_cover_preamble(line: str) -> bool:
    normalized = normalize_text(line)
    if not normalized:
        return True

    if normalized in {
        "viu",
        "universidad internacional de valencia",
        "universidad",
    }:
        return True

    return bool(
        re.fullmatch(
            r"\d{1,2}\s+[a-záéíóúü]{3,}(?:\s+de)?\s+\d{4}",
            normalized,
        )
        or re.fullmatch(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", normalized)
        or re.fullmatch(r"[a-záéíóúü]+,\s+[a-záéíóúü]+\s+\d{4}", normalized)
    )


def _collapse_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


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
