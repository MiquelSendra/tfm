"""Runtime configuration for the automation pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    """User-configurable runtime options."""

    workspace_dir: Path
    output_dir: Path
    match_threshold: float = 82.0
    ambiguity_margin: float = 4.0

    @classmethod
    def from_workspace(
        cls,
        workspace_dir: Path,
        output_dir_name: str = "output",
        match_threshold: float = 82.0,
        ambiguity_margin: float = 4.0,
    ) -> "AppConfig":
        """Build a config using paths rooted in the workspace directory."""
        workspace = workspace_dir.resolve()
        return cls(
            workspace_dir=workspace,
            output_dir=(workspace / output_dir_name).resolve(),
            match_threshold=match_threshold,
            ambiguity_margin=ambiguity_margin,
        )
