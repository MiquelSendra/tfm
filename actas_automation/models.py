"""Domain models used by the pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class ProgramMetadata:
    """Metadata extracted from the Excel file."""

    title: str = ""
    edition: str = ""


@dataclass(frozen=True)
class StudentRecord:
    """Student row extracted from the Excel roster."""

    full_name: str
    dni: str
    email: str
    director: str
    thesis_title: str
    thesis_topic: str
    source_row: int
    aliases: tuple[str, ...] = field(default_factory=tuple)

    @property
    def uid(self) -> str:
        """Stable identifier for matching and joining data sources."""
        if self.dni.strip():
            return f"dni:{self.dni.strip().lower()}"
        return f"name:{self.full_name.strip().lower()}"


@dataclass(frozen=True)
class SourceFiles:
    """Files discovered in the workspace."""

    excel_file: Path
    zip_file: Path
    docx_template: Optional[Path]
    pdf_files: list[Path]


@dataclass(frozen=True)
class SubmissionEntry:
    """A manuscript entry found inside the ZIP file."""

    zip_member: str
    extracted_name: str
    extension: str


@dataclass(frozen=True)
class DirectorReport:
    """Director report data parsed from a PDF."""

    path: Path
    extracted_name: str
    thesis_title: str
    extracted_text: str


@dataclass(frozen=True)
class NameMatch:
    """Result of fuzzy matching a candidate name against student records."""

    candidate_name: str
    matched_student: Optional[StudentRecord]
    score: float
    second_score: float
    status: str
    notes: str = ""


@dataclass(frozen=True)
class ActaContext:
    """Values inserted into the acta template."""

    titulacion: str
    edicion: str
    student_name: str
    dni: str
    thesis_title: str
    director: str
