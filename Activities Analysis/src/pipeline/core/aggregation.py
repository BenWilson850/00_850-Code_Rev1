"""Aggregation: summarize assessed results for reporting."""

from collections import defaultdict
from typing import Any

from ..utils import get_logger

_log = get_logger(__name__)


def aggregate_results(assessed: list[dict[str, Any]]) -> dict[str, Any]:
    """Aggregate assessed activity results into summary structures.

    Builds counts and lists per threshold and per activity for the report.

    Args:
        assessed: Output from assess_activities (rows with threshold_outcomes).

    Returns:
        Dict with keys such as 'by_threshold', 'by_activity', 'summary'.
    """
    by_threshold: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_activity: list[dict[str, Any]] = []
    summary: dict[str, int] = defaultdict(int)

    for row in assessed:
        activity_id = row.get("id") or row.get("name") or ""
        by_activity.append(row)
        for out in row.get("threshold_outcomes", []):
            tid = out.get("threshold_id") or "unknown"
            if out.get("met"):
                by_threshold[tid].append({"activity_id": activity_id, **row})
                summary[tid] += 1

    return {
        "by_threshold": dict(by_threshold),
        "by_activity": by_activity,
        "summary": dict(summary),
        "total_activities": len(assessed),
    }
