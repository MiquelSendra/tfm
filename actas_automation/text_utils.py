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


def stem_from_path(path: Path) -> str:
    """Return a clean stem from a path object."""
    return clean_text(path.stem)
