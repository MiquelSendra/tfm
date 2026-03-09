"""Pipeline orchestration for automatic acta generation."""

from __future__ import annotations

import csv
import logging
import os
import re
import subprocess
from shutil import which
from shutil import copy2
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .config import AppConfig
from .discovery import discover_source_files
from .docx_generator import generate_docx_acta
from .excel_source import load_students_and_metadata
from .matching import MatcherConfig, StudentMatcher
from .models import (
    ActaContext,
    DirectorReport,
    NameMatch,
    StudentRecord,
    SubmissionEntry,
)
from .reports import parse_pdf_files, try_fill_pdf_template
from .text_utils import (
    build_acta_filename,
    build_student_folder_name,
    clean_text,
    normalize_text,
)
from .zip_source import extract_zip_member, read_zip_submissions


@dataclass(frozen=True)
class PipelineOutcome:
    """High-level counters for final run reporting."""

    generated_actas: int
    manual_review_rows: int
    processed_submissions: int
    matched_submissions: int


@dataclass(frozen=True)
class _MatchedSubmission:
    student: StudentRecord
    match: NameMatch
    zip_member: str


def run_pipeline(config: AppConfig) -> PipelineOutcome:
    """Run the complete process and generate output artifacts."""
    output_dirs = _prepare_output_dirs(config.output_dir)
    logger = _build_logger(output_dirs["logs"])
    logger.info("Starting automation in workspace: %s", config.workspace_dir)

    sources = discover_source_files(config.workspace_dir, logger)
    metadata, students = load_students_and_metadata(str(sources.excel_file), logger)
    matcher = StudentMatcher(
        students=students,
        config=MatcherConfig(
            min_score=config.match_threshold,
            ambiguity_margin=config.ambiguity_margin,
        ),
    )

    submissions = read_zip_submissions(sources.zip_file, logger)
    matched_submissions, manual_rows = _match_submissions(submissions, matcher, logger)

    reports, non_report_pdfs = parse_pdf_files(sources.pdf_files, logger)
    report_by_student, report_strategy_by_student, report_manual_rows = _match_reports(
        reports,
        matcher,
        students,
        logger,
    )
    manual_rows.extend(report_manual_rows)

    summary_rows: list[dict[str, str]] = []
    generated_count = 0

    for student_uid, submission_info in matched_submissions.items():
        student = submission_info.student
        submission_match = submission_info.match
        submission_member = submission_info.zip_member

        report = report_by_student.get(student_uid)
        report_match_strategy = report_strategy_by_student.get(student_uid, "")
        context = _build_acta_context(student, metadata.title, metadata.edition, report)

        student_folder_name = build_student_folder_name(student.full_name)
        student_folder = output_dirs["students"] / student_folder_name
        student_folder.mkdir(parents=True, exist_ok=True)

        manuscript_name = _clean_submission_filename(Path(submission_member).name)
        manuscript_path = student_folder / manuscript_name
        legacy_manuscript_path = student_folder / Path(submission_member).name
        report_path_in_folder = ""
        report_notes = ""
        manuscript_status = "ok"
        manuscript_notes = ""

        try:
            if legacy_manuscript_path != manuscript_path and legacy_manuscript_path.exists():
                legacy_manuscript_path.unlink()
            extract_zip_member(sources.zip_file, submission_member, manuscript_path)
            manuscript_path = _convert_manuscript_to_pdf_if_needed(manuscript_path)
        except Exception as exc:  # pragma: no cover - depends on source files
            manuscript_status = "error"
            manuscript_notes = f"manuscript_processing_error: {exc}"
            logger.exception("Error processing manuscript for %s", student.full_name)

        if report:
            report_target = student_folder / report.path.name
            try:
                copy2(report.path, report_target)
                report_path_in_folder = str(report_target)
            except Exception as exc:  # pragma: no cover - depends on source files
                manuscript_status = "error"
                message = f"report_copy_error: {exc}"
                manuscript_notes = f"{manuscript_notes} | {message}" if manuscript_notes else message
                logger.exception("Error copying director report for %s", student.full_name)
        else:
            report_notes = "director_report_not_found"
            logger.warning("No director report matched for %s", student.full_name)

        output_name = build_acta_filename(student.full_name)
        acta_output_docx = student_folder / output_name
        generation_status = "ok"
        generation_notes = ""
        output_file = str(acta_output_docx)

        try:
            if sources.docx_template:
                generate_docx_acta(sources.docx_template, acta_output_docx, context)
            elif non_report_pdfs:
                pdf_output = student_folder / output_name.replace(".docx", ".pdf")
                success = try_fill_pdf_template(non_report_pdfs[0], pdf_output, context, logger)
                if not success:
                    raise RuntimeError("No suitable DOCX template or fillable PDF template found.")
                output_file = str(pdf_output)
            else:
                raise RuntimeError("No suitable DOCX template or fillable PDF template found.")
            generated_count += 1
        except Exception as exc:  # pragma: no cover - depends on source files
            generation_status = "error"
            generation_notes = str(exc)
            logger.exception("Error generating acta for %s", student.full_name)

        summary_rows.append(
            {
                "student_name": student.full_name,
                "student_folder": str(student_folder),
                "dni": student.dni,
                "zip_member": submission_member,
                "manuscript_output_file": str(manuscript_path),
                "zip_candidate_name": submission_match.candidate_name,
                "zip_match_score": f"{submission_match.score:.2f}",
                "director_report_file": report.path.name if report else "",
                "director_report_name": report.extracted_name if report else "",
                "director_report_match_strategy": report_match_strategy,
                "director_report_output_file": report_path_in_folder,
                "acta_output_file": output_file,
                "status": generation_status,
                "notes": " | ".join(
                    note for note in [generation_notes, manuscript_notes, report_notes] if note
                ),
                "manuscript_status": manuscript_status,
            }
        )

    summary_path = config.output_dir / "resumen_procesamiento.csv"
    manual_path = config.output_dir / "revision_manual.csv"
    _write_csv(
        summary_path,
        summary_rows,
        headers=[
            "student_name",
            "student_folder",
            "dni",
            "zip_member",
            "manuscript_output_file",
            "zip_candidate_name",
            "zip_match_score",
            "director_report_file",
            "director_report_name",
            "director_report_match_strategy",
            "director_report_output_file",
            "acta_output_file",
            "status",
            "manuscript_status",
            "notes",
        ],
    )
    _write_csv(
        manual_path,
        manual_rows,
        headers=[
            "source_type",
            "source_file",
            "candidate_name",
            "best_student",
            "best_score",
            "second_score",
            "issue",
            "notes",
        ],
    )

    logger.info("Actas generated: %s", generated_count)
    logger.info("Manual review rows: %s", len(manual_rows))
    logger.info("Summary CSV: %s", summary_path)
    logger.info("Manual review CSV: %s", manual_path)

    return PipelineOutcome(
        generated_actas=generated_count,
        manual_review_rows=len(manual_rows),
        processed_submissions=len(submissions),
        matched_submissions=len(matched_submissions),
    )


def _clean_submission_filename(filename: str) -> str:
    """Remove leading numeric prefix up to first dash from ZIP manuscript names."""
    cleaned = re.sub(r"^\s*\d+\s*-\s*", "", filename).strip()
    return cleaned or filename


def _convert_manuscript_to_pdf_if_needed(path: Path) -> Path:
    """Convert non-PDF manuscript files to PDF using LibreOffice headless."""
    if path.suffix.lower() == ".pdf":
        return path

    office_bin = _find_office_binary()
    if not office_bin:
        raise RuntimeError(
            "No LibreOffice/soffice binary found for manuscript conversion to PDF."
        )

    command = [
        office_bin,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(path.parent),
        str(path),
    ]
    result = subprocess.run(
        command,
        check=False,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            "LibreOffice conversion failed "
            f"(code {result.returncode}): {result.stderr.strip() or result.stdout.strip()}"
        )

    pdf_path = path.with_suffix(".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"Expected converted PDF not found: {pdf_path.name}")

    path.unlink(missing_ok=True)
    return pdf_path


def _find_office_binary() -> str:
    """Locate a usable LibreOffice executable."""
    env_bin = os.environ.get("LIBREOFFICE_BIN", "").strip()
    if env_bin and Path(env_bin).exists():
        return env_bin

    candidates = [
        "soffice",
        "libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for candidate in candidates:
        if candidate.startswith("/"):
            if Path(candidate).exists():
                return candidate
            continue
        found = which(candidate)
        if found:
            return found
    return ""


def _prepare_output_dirs(output_dir: Path) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    students_dir = output_dir / "estudiantes"
    logs_dir = output_dir / "logs"
    students_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    return {"base": output_dir, "students": students_dir, "logs": logs_dir}


def _build_logger(logs_dir: Path) -> logging.Logger:
    logger = logging.getLogger("actas_automation")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = logs_dir / f"procesamiento_{timestamp}.log"

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger


def _match_submissions(
    submissions: Iterable[SubmissionEntry],
    matcher: StudentMatcher,
    logger: logging.Logger,
) -> tuple[dict[str, _MatchedSubmission], list[dict[str, str]]]:
    matched_by_student: dict[str, _MatchedSubmission] = {}
    manual_rows: list[dict[str, str]] = []

    for entry in submissions:
        match = matcher.match(entry.extracted_name)
        if match.status == "matched" and match.matched_student:
            student_uid = match.matched_student.uid
            current = matched_by_student.get(student_uid)
            if current is None or match.score > current.match.score:
                matched_by_student[student_uid] = _MatchedSubmission(
                    student=match.matched_student,
                    match=match,
                    zip_member=entry.zip_member,
                )
            else:
                manual_rows.append(
                    _manual_row(
                        source_type="zip_duplicate",
                        source_file=entry.zip_member,
                        match=match,
                        issue="duplicate_submission",
                    )
                )
        else:
            manual_rows.append(
                _manual_row(
                    source_type="zip",
                    source_file=entry.zip_member,
                    match=match,
                    issue=match.status,
                )
            )
            logger.warning(
                "Submission requires manual review: '%s' (%s, %.2f)",
                entry.extracted_name,
                match.status,
                match.score,
            )

    return matched_by_student, manual_rows


def _match_reports(
    reports: list[DirectorReport],
    matcher: StudentMatcher,
    students: list[StudentRecord],
    logger: logging.Logger,
) -> tuple[dict[str, DirectorReport], dict[str, str], list[dict[str, str]]]:
    report_by_student: dict[str, DirectorReport] = {}
    report_strategy_by_student: dict[str, str] = {}
    manual_rows: list[dict[str, str]] = []
    surname_index = _build_surname_index(students)

    for report in reports:
        match = matcher.match(report.extracted_name)
        if match.status == "matched" and match.matched_student:
            _assign_report_match(
                report_by_student,
                report_strategy_by_student,
                manual_rows,
                report,
                match.matched_student,
                "full_name",
                logger,
            )
            continue

        surname_key = _extract_surname_key(report.extracted_name)
        if surname_key:
            candidates = surname_index.get(surname_key, [])
            if len(candidates) == 1:
                student = candidates[0]
                _assign_report_match(
                    report_by_student,
                    report_strategy_by_student,
                    manual_rows,
                    report,
                    student,
                    "surname_fallback",
                    logger,
                )
                continue
            if len(candidates) > 1:
                manual_rows.append(
                    {
                        "source_type": "director_report",
                        "source_file": report.path.name,
                        "candidate_name": report.extracted_name,
                        "best_student": "",
                        "best_score": f"{match.score:.2f}",
                        "second_score": f"{match.second_score:.2f}",
                        "issue": "surname_ambiguous",
                        "notes": ";".join(student.full_name for student in candidates),
                    }
                )
                logger.warning(
                    "Ambiguous surname fallback for report %s (%s candidates).",
                    report.path.name,
                    len(candidates),
                )
                continue

        if match.status != "matched":
            manual_rows.append(
                _manual_row(
                    source_type="director_report",
                    source_file=report.path.name,
                    match=match,
                    issue=match.status,
                )
            )

    return report_by_student, report_strategy_by_student, manual_rows


def _assign_report_match(
    report_by_student: dict[str, DirectorReport],
    report_strategy_by_student: dict[str, str],
    manual_rows: list[dict[str, str]],
    report: DirectorReport,
    student: StudentRecord,
    strategy: str,
    logger: logging.Logger,
) -> None:
    existing_report = report_by_student.get(student.uid)
    existing_strategy = report_strategy_by_student.get(student.uid, "")

    if not existing_report:
        report_by_student[student.uid] = report
        report_strategy_by_student[student.uid] = strategy
        return

    # Prefer full-name matches over surname fallback when duplicates exist.
    if existing_strategy == "surname_fallback" and strategy == "full_name":
        manual_rows.append(
            {
                "source_type": "director_report_duplicate",
                "source_file": existing_report.path.name,
                "candidate_name": existing_report.extracted_name,
                "best_student": student.full_name,
                "best_score": "0.00",
                "second_score": "0.00",
                "issue": "replaced_by_full_name_match",
                "notes": report.path.name,
            }
        )
        report_by_student[student.uid] = report
        report_strategy_by_student[student.uid] = strategy
        return

    manual_rows.append(
        {
            "source_type": "director_report_duplicate",
            "source_file": report.path.name,
            "candidate_name": report.extracted_name,
            "best_student": student.full_name,
            "best_score": "0.00",
            "second_score": "0.00",
            "issue": "duplicate_report_for_student",
            "notes": f"existing={existing_report.path.name};strategy={existing_strategy}",
        }
    )
    logger.warning(
        "Duplicate report for student %s ignored: %s",
        student.full_name,
        report.path.name,
    )


def _build_surname_index(students: list[StudentRecord]) -> dict[str, list[StudentRecord]]:
    index: dict[str, list[StudentRecord]] = {}
    for student in students:
        key = _extract_surname_key(student.full_name)
        if not key:
            continue
        index.setdefault(key, []).append(student)
    return index


def _extract_surname_key(name: str) -> str:
    value = clean_text(name)
    if not value:
        return ""
    if "," in value:
        surname = value.split(",", 1)[0].strip()
    else:
        surname = value.strip()
    return normalize_text(surname)


def _build_acta_context(
    student: StudentRecord,
    titulacion: str,
    edicion: str,
    report: DirectorReport | None,
) -> ActaContext:
    thesis_title = clean_text(student.thesis_title)
    if not thesis_title and report:
        thesis_title = clean_text(report.thesis_title)
    if not thesis_title:
        thesis_title = clean_text(student.thesis_topic)

    return ActaContext(
        titulacion=titulacion,
        edicion=edicion,
        student_name=student.full_name,
        dni=student.dni,
        thesis_title=thesis_title,
        director=student.director,
    )


def _manual_row(
    source_type: str,
    source_file: str,
    match: NameMatch,
    issue: str,
) -> dict[str, str]:
    best_student = match.matched_student.full_name if match.matched_student else ""
    return {
        "source_type": source_type,
        "source_file": source_file,
        "candidate_name": match.candidate_name,
        "best_student": best_student,
        "best_score": f"{match.score:.2f}",
        "second_score": f"{match.second_score:.2f}",
        "issue": issue,
        "notes": match.notes,
    }


def _write_csv(path: Path, rows: list[dict[str, str]], headers: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
