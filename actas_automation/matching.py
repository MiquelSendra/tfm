"""Name matching utilities based on fuzzy similarity."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

from rapidfuzz import fuzz

from .models import NameMatch, StudentRecord
from .text_utils import normalize_text


@dataclass(frozen=True)
class MatcherConfig:
    """Thresholds controlling fuzzy matching behavior."""

    min_score: float
    ambiguity_margin: float


class StudentMatcher:
    """Resolve candidate names to known students with ambiguity handling."""

    def __init__(self, students: Iterable[StudentRecord], config: MatcherConfig):
        self._students = list(students)
        self._config = config

    def match(self, candidate_name: str) -> NameMatch:
        """Return match status for one candidate name."""
        candidate_norm = normalize_text(candidate_name)
        if not candidate_norm:
            return NameMatch(
                candidate_name=candidate_name,
                matched_student=None,
                score=0.0,
                second_score=0.0,
                status="unmatched",
                notes="empty_candidate",
            )

        scores: list[tuple[float, StudentRecord]] = []
        for student in self._students:
            aliases = student.aliases or (student.full_name,)
            alias_scores = [
                fuzz.token_sort_ratio(candidate_norm, normalize_text(alias))
                for alias in aliases
            ]
            best_alias_score = max(alias_scores) if alias_scores else 0.0
            scores.append((best_alias_score, student))

        if not scores:
            return NameMatch(
                candidate_name=candidate_name,
                matched_student=None,
                score=0.0,
                second_score=0.0,
                status="unmatched",
                notes="no_students_loaded",
            )

        scores.sort(key=lambda item: item[0], reverse=True)
        top_score, top_student = scores[0]
        second_score = scores[1][0] if len(scores) > 1 else 0.0

        if top_score < self._config.min_score:
            return NameMatch(
                candidate_name=candidate_name,
                matched_student=None,
                score=top_score,
                second_score=second_score,
                status="unmatched",
                notes="score_below_threshold",
            )

        if (
            second_score >= self._config.min_score
            and (top_score - second_score) < self._config.ambiguity_margin
        ):
            dni_tie_break = self._resolve_tie_by_dni(scores, top_score)
            if dni_tie_break:
                return NameMatch(
                    candidate_name=candidate_name,
                    matched_student=dni_tie_break,
                    score=top_score,
                    second_score=second_score,
                    status="matched",
                    notes="tie_break_by_dni",
                )
            return NameMatch(
                candidate_name=candidate_name,
                matched_student=top_student,
                score=top_score,
                second_score=second_score,
                status="ambiguous",
                notes="top_two_too_close",
            )

        return NameMatch(
            candidate_name=candidate_name,
            matched_student=top_student,
            score=top_score,
            second_score=second_score,
            status="matched",
            notes="ok",
        )

    @staticmethod
    def _resolve_tie_by_dni(
        scores: list[tuple[float, StudentRecord]],
        top_score: float,
    ) -> StudentRecord | None:
        """Prefer the only tied candidate that has a DNI."""
        tied_students = [
            student
            for score, student in scores
            if abs(score - top_score) < 1e-9
        ]
        if len(tied_students) <= 1:
            return None

        with_dni = [student for student in tied_students if student.dni.strip()]
        if len(with_dni) == 1:
            return with_dni[0]
        return None
