"""CLI entrypoint for TFT/TFM acta generation."""

from __future__ import annotations

import argparse
from pathlib import Path

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
        default=Path.cwd(),
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
    return parser


def main() -> int:
    """Run the automation pipeline."""
    args = build_parser().parse_args()
    config = AppConfig.from_workspace(
        workspace_dir=args.workspace,
        output_dir_name=args.output_dir_name,
        match_threshold=args.match_threshold,
        ambiguity_margin=args.ambiguity_margin,
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
