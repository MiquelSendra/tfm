"""Text normalization and naming helpers."""

from __future__ import annotations

import re
from pathlib import Path

from unidecode import unidecode


def normalize_text(value: str) -> str:
    """Normalize text for resilient matching across sources."""
    value = unidecode(value or "")
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def clean_text(value: object) -> str:
    """Convert any scalar into a stripped string."""
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def build_name_aliases(name: str) -> tuple[str, ...]:
    """Create common name variants to improve fuzzy matching quality."""
    cleaned = " ".join(clean_text(name).split())
    if not cleaned:
        return tuple()

    aliases: set[str] = {cleaned}

    if "," in cleaned:
        surname, given = [part.strip() for part in cleaned.split(",", 1)]
        if given and surname:
            aliases.add(f"{given} {surname}")
            aliases.add(f"{surname} {given}")
    else:
        parts = cleaned.split()
        if len(parts) >= 3:
            given = " ".join(parts[:-2])
            surnames = " ".join(parts[-2:])
            aliases.add(f"{surnames}, {given}")
            aliases.add(f"{surnames} {given}")

    return tuple(sorted(aliases))


def build_acta_filename(student_name: str, suffix: str = "acta.docx") -> str:
    """Build a safe and consistent output filename for the generated acta."""
    name = clean_text(student_name)
    if "," in name:
        surname, given = [part.strip() for part in name.split(",", 1)]
        joined = f"{surname}_{given}"
    else:
        joined = name.replace(" ", "_")

    ascii_name = unidecode(joined)
    ascii_name = re.sub(r"[^A-Za-z0-9_]+", "_", ascii_name)
    ascii_name = re.sub(r"_+", "_", ascii_name).strip("_")
    return f"{ascii_name}_{suffix}"


def extract_template_code_and_edition(template_path: Path | None) -> tuple[str, str]:
    """Extract subject code and edition from '<code>_<name>_<edition>' stems."""
    if not template_path:
        return "", ""

    stem = clean_text(template_path.stem)
    if not stem:
        return "", ""

    parts = stem.split("_")
    if not parts:
        return "", ""

    code = clean_text(parts[0])
    edition = clean_text("_".join(parts[2:])) if len(parts) >= 3 else ""
    return code, edition


def build_acta_output_stem(
    student_name: str,
    template_path: Path | None = None,
    edition: str = "",
) -> str:
    """Build acta stem preserving template naming convention when available."""
    student_display = build_student_folder_name(student_name)
    student_display = re.sub(r"\s+", " ", student_display).strip()
    student_display = student_display.replace("/", "-")

    code, template_edition = extract_template_code_and_edition(template_path)
    resolved_edition = clean_text(edition) or template_edition

    if code:
        parts = [code, student_display]
        if resolved_edition:
            parts.append(resolved_edition)
        return "_".join(parts)

    # Fallback for non-conforming template names.
    if template_path:
        stem = clean_text(template_path.stem)
        if stem:
            return f"{stem}_{student_display}"

    return build_acta_filename(student_display, suffix="acta").removesuffix("_acta")


def build_student_folder_name(student_name: str) -> str:
    """Build display folder name as 'Apellido1 Apellido2, Nombre'."""
    name = " ".join(clean_text(student_name).split())
    if not name:
        return "Estudiante_sin_nombre"

    if "," in name:
        surname, given = [part.strip() for part in name.split(",", 1)]
        display_name = f"{surname}, {given}"
    else:
        parts = name.split()
        if len(parts) >= 3:
            display_name = f"{' '.join(parts[-2:])}, {' '.join(parts[:-2])}"
        else:
            display_name = name

    display_name = re.sub(r"\s*,\s*", ", ", display_name).strip()
    display_name = display_name.replace("/", "-")
    return display_name


def build_student_document_name(prefix: str, student_name: str, extension: str = ".pdf") -> str:
    """Build a descriptive ASCII filename like 'prefix_Apellido_Apellido_Nombre.pdf'."""
    display_name = build_student_folder_name(student_name)
    ascii_name = unidecode(display_name)
    ascii_name = ascii_name.replace(", ", "_")
    ascii_name = ascii_name.replace(",", "_")
    ascii_name = re.sub(r"[^A-Za-z0-9_]+", "_", ascii_name)
    ascii_name = re.sub(r"_+", "_", ascii_name).strip("_")
    normalized_prefix = re.sub(r"[^A-Za-z0-9_]+", "_", clean_text(prefix))
    normalized_prefix = re.sub(r"_+", "_", normalized_prefix).strip("_")
    normalized_extension = extension if extension.startswith(".") else f".{extension}"
    return f"{normalized_prefix}_{ascii_name}{normalized_extension}"


def stem_from_path(path: Path) -> str:
    """Return a clean stem from a path object."""
    return clean_text(path.stem)
