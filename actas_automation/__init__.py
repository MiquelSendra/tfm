"""Automation package for generating TFT/TFM evaluation reports."""

from .config import AppConfig
from .pipeline import PipelineOutcome, run_pipeline

__all__ = ["AppConfig", "PipelineOutcome", "run_pipeline"]
