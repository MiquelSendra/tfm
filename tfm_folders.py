"""CLI entrypoint for TFT/TFM acta generation."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

from actas_automation import AppConfig, run_pipeline


def build_parser() -> argparse.ArgumentParser:
    """Create CLI arguments parser."""
    parser = argparse.ArgumentParser(
        description=(
            "Generate evaluation reports (actas) for students with manuscripts in ZIP."
        )
    )
    parser.add_argument(
        "--workspace",
        type=Path,
        default=None,
        help="Workspace directory containing ZIP, Excel, templates and reports.",
    )
    parser.add_argument(
        "--output-dir-name",
        default="output",
        help="Output directory name created under --workspace.",
    )
    parser.add_argument(
        "--match-threshold",
        type=float,
        default=82.0,
        help="Minimum fuzzy match score (0-100) to accept automatic matches.",
    )
    parser.add_argument(
        "--ambiguity-margin",
        type=float,
        default=4.0,
        help="Minimum gap between best and second match to avoid ambiguity.",
    )
    parser.add_argument(
        "--only-student",
        default="",
        help=(
            "Optional substring filter to process only matching student names "
            "(case/accent insensitive)."
        ),
    )
    return parser


def _resolve_workspace(workspace_arg: Path | None) -> Path:
    """Resolve default workspace, favoring obvious run folders."""
    if workspace_arg:
        return workspace_arg

    cwd = Path.cwd()
    executable_dir = Path(sys.executable).resolve().parent

    if _looks_like_workspace(cwd):
        return cwd
    if _looks_like_workspace(executable_dir):
        return executable_dir
    if getattr(sys, "frozen", False):
        return executable_dir
    return cwd


def _looks_like_workspace(path: Path) -> bool:
    return any(path.glob("*.zip")) and any(path.glob("*.xlsx"))


def main() -> int:
    """Run the automation pipeline."""
    args = build_parser().parse_args()
    workspace = _resolve_workspace(args.workspace)
    config = AppConfig.from_workspace(
        workspace_dir=workspace,
        output_dir_name=args.output_dir_name,
        match_threshold=args.match_threshold,
        ambiguity_margin=args.ambiguity_margin,
        student_filter=args.only_student,
    )
    outcome = run_pipeline(config)
    print(
        "Proceso completado. "
        f"Actas generadas: {outcome.generated_actas} | "
        f"Envíos ZIP procesados: {outcome.processed_submissions} | "
        f"Envíos ZIP emparejados: {outcome.matched_submissions} | "
        f"Filas para revisión manual: {outcome.manual_review_rows}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
