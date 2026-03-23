"""Pipeline orchestration for automatic acta generation."""

from __future__ import annotations

import csv
import logging
import re
from shutil import copy2
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .config import AppConfig
from .discovery import discover_source_files
from .excel_source import load_students_and_metadata
from .matching import MatcherConfig, StudentMatcher
from .models import (
    ActaContext,
    DirectorReport,
    NameMatch,
    StudentRecord,
    SubmissionEntry,
)
from .reports import extract_manuscript_title_from_pdf, parse_pdf_files, try_fill_pdf_template
from .supporting_documents import (
    SupportingDocument,
    copy_supporting_document,
    read_slide_documents,
)
from .text_utils import (
    build_acta_output_stem,
    build_student_document_name,
    build_student_folder_name,
    clean_text,
    extract_template_code_and_edition,
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


@dataclass(frozen=True)
class _MatchedSupportingDocument:
    student: StudentRecord
    document: SupportingDocument
    strategy: str


def run_pipeline(config: AppConfig) -> PipelineOutcome:
    """Run the complete process and generate output artifacts."""
    output_dirs = _prepare_output_dirs(config.output_dir)
    logger = _build_logger(output_dirs["logs"])
    logging.getLogger("pypdf").setLevel(logging.ERROR)
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
    matched_submissions = _filter_matched_submissions(
        matched_submissions,
        config.student_filter,
        logger,
    )

    reports, _ = parse_pdf_files(sources.pdf_files, logger)
    report_by_student, report_strategy_by_student, report_manual_rows = _match_reports(
        reports,
        matcher,
        students,
        logger,
    )
    manual_rows.extend(report_manual_rows)
    slides = read_slide_documents(config.workspace_dir / "diapositivas", logger)
    slide_by_student, slide_manual_rows = _match_supporting_documents(
        slides,
        matcher,
        students,
        logger,
        source_type="slides",
    )
    manual_rows.extend(slide_manual_rows)

    summary_rows: list[dict[str, str]] = []
    generated_count = 0
    _, template_edition = extract_template_code_and_edition(sources.acta_pdf_template)
    resolved_edition = clean_text(metadata.edition) or template_edition
    if not clean_text(metadata.edition) and resolved_edition:
        logger.info(
            "Edition inferred from template filename: '%s'", resolved_edition
        )

    for student_uid, submission_info in matched_submissions.items():
        student = submission_info.student
        submission_match = submission_info.match
        submission_member = submission_info.zip_member

        report = report_by_student.get(student_uid)
        report_match_strategy = report_strategy_by_student.get(student_uid, "")
        slide_document = slide_by_student.get(student_uid)

        student_folder_name = build_student_folder_name(student.full_name)
        student_folder = output_dirs["students"] / student_folder_name
        student_folder.mkdir(parents=True, exist_ok=True)
        output_stem = build_acta_output_stem(
            student_name=student.full_name,
            template_path=sources.acta_pdf_template,
            edition=resolved_edition,
        )
        acta_output_pdf = student_folder / f"{output_stem}.pdf"
        output_file = str(acta_output_pdf)

        manuscript_name = _clean_submission_filename(Path(submission_member).name)
        manuscript_path = student_folder / manuscript_name
        legacy_manuscript_path = student_folder / Path(submission_member).name
        report_path_in_folder = ""
        report_status = ""
        report_notes = ""
        manuscript_status = "ok"
        manuscript_notes = ""
        manuscript_title = ""
        slide_path_in_folder = ""
        slide_source_file = ""
        slide_status = ""
        slide_notes = ""

        try:
            existing_manuscript = _resolve_existing_manuscript_path(
                manuscript_path,
                legacy_manuscript_path,
            )
            if existing_manuscript:
                manuscript_path = existing_manuscript
                manuscript_status = "existing"
                logger.info(
                    "Skipping manuscript extraction for %s because it already exists.",
                    student.full_name,
                )
                manuscript_conversion_note = ""
            else:
                if legacy_manuscript_path != manuscript_path and legacy_manuscript_path.exists():
                    legacy_manuscript_path.unlink()
                extract_zip_member(sources.zip_file, submission_member, manuscript_path)
                manuscript_path, manuscript_conversion_note = _convert_manuscript_to_pdf_if_needed(
                    manuscript_path
                )
            if manuscript_path.exists() and manuscript_path.suffix.lower() == ".pdf":
                manuscript_title = extract_manuscript_title_from_pdf(manuscript_path, logger)
                if manuscript_title:
                    logger.info(
                        "Manuscript title extracted for %s: %s",
                        student.full_name,
                        manuscript_title,
                    )
                else:
                    message = "manuscript_title_not_found"
                    manuscript_notes = (
                        f"{manuscript_notes} | {message}" if manuscript_notes else message
                    )
            if manuscript_conversion_note:
                manuscript_notes = (
                    f"{manuscript_notes} | {manuscript_conversion_note}"
                    if manuscript_notes
                    else manuscript_conversion_note
                )
        except Exception as exc:  # pragma: no cover - depends on source files
            manuscript_status = "error"
            manuscript_notes = f"manuscript_processing_error: {exc}"
            logger.exception("Error processing manuscript for %s", student.full_name)

        report_target = student_folder / build_student_document_name(
            "informe_director",
            student.full_name,
        )
        legacy_report_target = student_folder / "informe_director.pdf"
        if report:
            try:
                if _promote_legacy_named_file(legacy_report_target, report_target):
                    report_status = "existing"
                elif report_target.exists():
                    report_status = "existing"
                else:
                    copy2(report.path, report_target)
                    report_status = "copied"
                report_path_in_folder = str(report_target)
            except Exception as exc:  # pragma: no cover - depends on source files
                report_status = "error"
                message = f"report_copy_error: {exc}"
                report_notes = f"{report_notes} | {message}" if report_notes else message
                logger.exception("Error copying director report for %s", student.full_name)
        elif _promote_legacy_named_file(legacy_report_target, report_target) or report_target.exists():
            report_status = "existing"
            report_path_in_folder = str(report_target)
        else:
            report_status = "missing"
            report_notes = "director_report_not_found"
            logger.warning("No director report matched for %s", student.full_name)

        slide_target = student_folder / build_student_document_name(
            "diapositivas",
            student.full_name,
        )
        legacy_slide_target = student_folder / "diapositivas.pdf"
        if slide_document:
            try:
                if _promote_legacy_named_file(legacy_slide_target, slide_target):
                    slide_status = "existing"
                elif slide_target.exists():
                    slide_status = "existing"
                else:
                    copy_supporting_document(slide_document, slide_target)
                    slide_status = "copied"
                slide_path_in_folder = str(slide_target)
                slide_source_file = slide_document.source_file
            except Exception as exc:  # pragma: no cover - depends on source files
                slide_status = "error"
                slide_notes = f"slides_copy_error: {exc}"
                logger.exception("Error copying slides for %s", student.full_name)
        elif _promote_legacy_named_file(legacy_slide_target, slide_target) or slide_target.exists():
            slide_status = "existing"
            slide_path_in_folder = str(slide_target)
        else:
            slide_status = "missing"

        context = _build_acta_context(
            student,
            metadata.title,
            resolved_edition,
            report,
            manuscript_title,
        )
        generation_status = "created"
        generation_notes = ""

        if acta_output_pdf.exists():
            generation_status = "existing"
            logger.info(
                "Skipping acta generation for %s because it already exists.",
                student.full_name,
            )
        else:
            try:
                success = try_fill_pdf_template(
                    sources.acta_pdf_template,
                    acta_output_pdf,
                    context,
                    logger,
                )
                if not success:
                    raise RuntimeError("Acta PDF template has no fillable fields.")
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
                "director_report_status": report_status,
                "director_report_output_file": report_path_in_folder,
                "slides_status": slide_status,
                "slides_source_file": slide_source_file,
                "slides_output_file": slide_path_in_folder,
                "acta_output_file": output_file,
                "status": generation_status,
                "notes": " | ".join(
                    note
                    for note in [
                        generation_notes,
                        manuscript_notes,
                        report_notes,
                        slide_notes,
                    ]
                    if note
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
            "director_report_status",
            "director_report_output_file",
            "slides_status",
            "slides_source_file",
            "slides_output_file",
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


def _convert_manuscript_to_pdf_if_needed(path: Path) -> tuple[Path, str]:
    """Keep manuscript as-is; no office conversion dependency is required."""
    if path.suffix.lower() == ".pdf":
        return path, ""
    return path, "manuscript_kept_original_non_pdf"


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


def _filter_matched_submissions(
    matched_by_student: dict[str, _MatchedSubmission],
    student_filter: str,
    logger: logging.Logger,
) -> dict[str, _MatchedSubmission]:
    if not clean_text(student_filter):
        return matched_by_student

    token = normalize_text(student_filter)
    filtered = {
        uid: info
        for uid, info in matched_by_student.items()
        if token in normalize_text(info.student.full_name)
    }
    logger.info(
        "Applied --only-student filter '%s': %s -> %s matches",
        student_filter,
        len(matched_by_student),
        len(filtered),
    )
    return filtered


def _match_supporting_documents(
    documents: list[SupportingDocument],
    matcher: StudentMatcher,
    students: list[StudentRecord],
    logger: logging.Logger,
    source_type: str,
) -> tuple[dict[str, SupportingDocument], list[dict[str, str]]]:
    matched_by_student: dict[str, _MatchedSupportingDocument] = {}
    manual_rows: list[dict[str, str]] = []
    email_index = _build_email_index(students)

    for document in documents:
        matched_student: StudentRecord | None = None
        strategy = ""

        if document.email and document.email in email_index:
            matched_student = email_index[document.email]
            strategy = "email"
        elif document.candidate_name:
            match = matcher.match(document.candidate_name)
            if match.status == "matched" and match.matched_student:
                matched_student = match.matched_student
                strategy = "cover_name"
            else:
                manual_rows.append(
                    _manual_row(
                        source_type=source_type,
                        source_file=document.source_file,
                        match=match,
                        issue=match.status,
                    )
                )
                logger.warning(
                    "%s requires manual review: '%s' (%s, %.2f)",
                    source_type,
                    document.candidate_name,
                    match.status,
                    match.score,
                )
                continue
        else:
            manual_rows.append(
                {
                    "source_type": source_type,
                    "source_file": document.source_file,
                    "candidate_name": "",
                    "best_student": "",
                    "best_score": "0.00",
                    "second_score": "0.00",
                    "issue": "name_not_found",
                    "notes": document.email,
                }
            )
            logger.warning("%s skipped without candidate name: %s", source_type, document.source_file)
            continue

        existing = matched_by_student.get(matched_student.uid)
        if not existing:
            matched_by_student[matched_student.uid] = _MatchedSupportingDocument(
                student=matched_student,
                document=document,
                strategy=strategy,
            )
            continue

        if existing.strategy != "email" and strategy == "email":
            manual_rows.append(
                {
                    "source_type": f"{source_type}_duplicate",
                    "source_file": existing.document.source_file,
                    "candidate_name": existing.document.candidate_name,
                    "best_student": matched_student.full_name,
                    "best_score": "0.00",
                    "second_score": "0.00",
                    "issue": "replaced_by_email_match",
                    "notes": document.source_file,
                }
            )
            matched_by_student[matched_student.uid] = _MatchedSupportingDocument(
                student=matched_student,
                document=document,
                strategy=strategy,
            )
            continue

        manual_rows.append(
            {
                "source_type": f"{source_type}_duplicate",
                "source_file": document.source_file,
                "candidate_name": document.candidate_name,
                "best_student": matched_student.full_name,
                "best_score": "0.00",
                "second_score": "0.00",
                "issue": f"duplicate_{source_type}_for_student",
                "notes": (
                    f"existing={existing.document.source_file};"
                    f"strategy={existing.strategy}"
                ),
            }
        )
        logger.warning(
            "Duplicate %s for student %s ignored: %s",
            source_type,
            matched_student.full_name,
            document.source_file,
        )

    return {
        student_uid: item.document
        for student_uid, item in matched_by_student.items()
    }, manual_rows


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


def _build_email_index(students: list[StudentRecord]) -> dict[str, StudentRecord]:
    index: dict[str, StudentRecord] = {}
    for student in students:
        email = clean_text(student.email).lower()
        if email and email not in index:
            index[email] = student
    return index


def _resolve_existing_manuscript_path(
    manuscript_path: Path,
    legacy_manuscript_path: Path,
) -> Path | None:
    for candidate in (manuscript_path, legacy_manuscript_path):
        if candidate.exists():
            return candidate
    return None


def _promote_legacy_named_file(legacy_path: Path, target_path: Path) -> bool:
    if target_path.exists():
        return True
    if not legacy_path.exists():
        return False
    legacy_path.replace(target_path)
    return True


def _build_acta_context(
    student: StudentRecord,
    titulacion: str,
    edicion: str,
    report: DirectorReport | None,
    manuscript_title: str,
) -> ActaContext:
    thesis_title = clean_text(manuscript_title)
    if not thesis_title and report:
        thesis_title = clean_text(report.thesis_title)
    if not thesis_title:
        thesis_title = clean_text(student.thesis_title)
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
