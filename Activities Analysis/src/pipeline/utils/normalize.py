"""Normalization: convert DataFrame rows to consistent dicts for the pipeline."""

import pandas as pd
from typing import Any


def normalize_activity_row(row: pd.Series) -> dict[str, Any]:
    """Normalize a single activity row from the activity DataFrame.

    Strips column names, drops NaN, and uses first column as id/name if present.
    """
    d = row.to_dict()
    out = {}
    for k, v in d.items():
        key = str(k).strip() if k is not None else ""
        if pd.isna(v):
            continue
        out[key] = v
    # Prefer id/name for display
    if "id" not in out and "name" not in out and out:
        first_key = next(iter(out))
        out["name"] = out.get("name", out[first_key])
    return out


def normalize_threshold_row(row: pd.Series) -> dict[str, Any]:
    """Normalize a single threshold row from the threshold matrix DataFrame.

    Produces a dict with consistent keys (id, name, min_value, etc.) for assessment.
    """
    d = row.to_dict()
    out = {}
    for k, v in d.items():
        key = str(k).strip().lower().replace(" ", "_") if k is not None else ""
        if pd.isna(v):
            continue
        out[key] = v
    if "id" not in out:
        out["id"] = out.get("name") or out.get("threshold") or str(list(out.values())[0]) if out else ""
    return out
