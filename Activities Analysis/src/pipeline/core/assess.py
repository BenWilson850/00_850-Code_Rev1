"""Assessment: compare activity data against threshold matrix logic."""

from typing import Any, Literal

from ..utils import get_logger

_log = get_logger(__name__)


def assess_activities(
    activity_rows: list[dict[str, Any]],
    threshold_rows: list[dict[str, Any]],
    *,
    mode: Literal["strict", "inclusive"] = "inclusive",
) -> list[dict[str, Any]]:
    """Assess each activity row against the threshold matrix.

    For each activity, determines which thresholds are met (or exceeded)
    according to the configured mode.

    Args:
        activity_rows: Normalized activity records.
        threshold_rows: Normalized threshold/classification rows.
        mode: 'strict' = must exceed; 'inclusive' = meet or exceed.

    Returns:
        List of assessed records with threshold outcomes attached.
    """
    results = []
    for row in activity_rows:
        outcomes = []
        for th in threshold_rows:
            met = _evaluate_threshold(row, th, mode)
            outcomes.append({"threshold_id": th.get("id"), "met": met})
        results.append({**row, "threshold_outcomes": outcomes})
    return results


def _evaluate_threshold(
    activity: dict[str, Any],
    threshold: dict[str, Any],
    mode: str,
) -> bool:
    """Evaluate a single threshold against an activity row."""
    # Placeholder: compare activity fields to threshold criteria.
    # Extend with real column names from your matrix (e.g. intensity, duration).
    threshold_id = threshold.get("id") or threshold.get("name")
    if not threshold_id:
        return False
    # Example: if threshold has 'min_value' and activity has 'value'
    min_val = threshold.get("min_value")
    if min_val is None:
        return True
    act_val = activity.get("value") or activity.get("intensity")
    if act_val is None:
        return False
    try:
        if mode == "inclusive":
            return float(act_val) >= float(min_val)
        return float(act_val) > float(min_val)
    except (TypeError, ValueError):
        return False
