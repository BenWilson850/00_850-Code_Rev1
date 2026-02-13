"""Core parsing and scoring logic for the persona workbook matrix pipeline."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Literal

import pandas as pd

from .persona_models import LimitSpec, Importance, Zone


_OP_RE = re.compile(r"(<=|>=|<|>)")
_NUM_RE = re.compile(r"(-?\d+(?:\.\d+)?)")


@dataclass(frozen=True)
class AggregationRules:
    supporting_red_for_red: int  # strictly greater than this triggers RED
    supporting_red_for_yellow: int  # exactly this triggers YELLOW


DEFAULT_RULES = AggregationRules(supporting_red_for_red=3, supporting_red_for_yellow=2)


def infer_importance(cell: Any) -> Importance | None:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return None
    text = str(cell).lower()
    if "critical" in text:
        return "Critical"
    if "supporting" in text:
        return "Supporting"
    return None


def clean_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not (isinstance(value, float) and pd.isna(value)):
        return float(value)
    s = str(value).strip()
    if not s or s.lower() in {"na", "n/a", "nan", "none"}:
        return None
    s = s.replace("%", "")
    m = _NUM_RE.search(s)
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


def parse_limit(raw: Any, *, gender: Literal["M", "F"] | None) -> LimitSpec | None:
    """Parse a limit cell like '>15', '<9.0%', '>15 (F), >20 (M)'."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    if isinstance(raw, (int, float)) and not (isinstance(raw, float) and pd.isna(raw)):
        # Best-effort: a bare number is treated as >= threshold.
        return LimitSpec(op=">=", value=float(raw))

    s = str(raw).strip()
    if not s:
        return None

    parts = [p.strip() for p in s.split(",") if p.strip()]
    if not parts:
        parts = [s]

    def _part_gender(p: str) -> Literal["M", "F"] | None:
        pl = p.lower()
        if "(m)" in pl or " male" in pl:
            return "M"
        if "(f)" in pl or " female" in pl:
            return "F"
        return None

    def _parse_one(p: str) -> LimitSpec | None:
        opm = _OP_RE.search(p)
        if not opm:
            return None
        op_str = opm.group(1)
        # Ensure op is one of the valid literal types
        if op_str not in (">", ">=", "<", "<="):
            return None
        op: Literal[">", ">=", "<", "<="] = op_str  # type: ignore[assignment]
        num = clean_number(p)
        if num is None:
            return None
        return LimitSpec(op=op, value=num)

    parsed: list[tuple[Literal["M", "F"] | None, LimitSpec]] = []
    for p in parts:
        spec = _parse_one(p)
        if spec:
            parsed.append((_part_gender(p), spec))
    if not parsed:
        # Try parsing the whole string
        spec = _parse_one(s)
        return spec

    # If a matching gender exists, use it.
    if gender:
        for g, spec in parsed:
            if g == gender:
                return spec

    # Gender unknown or no match: default to the "easier" threshold to avoid false failures.
    # For direct metrics (>=), smaller is easier. For inverse metrics (<=), larger is easier.
    direct = any(spec.op in {">", ">="} for _, spec in parsed)
    if direct:
        return min((spec for _, spec in parsed), key=lambda x: x.value)
    return max((spec for _, spec in parsed), key=lambda x: x.value)


def compute_pi(client_value: float | None, limit: LimitSpec) -> float | None:
    if client_value is None:
        return None
    limit_value = limit.value
    op = limit.op

    # Default formula per spec, with a robust fallback for negatives/zeros.
    direct = op in {">", ">="}
    if direct and limit_value > 0 and client_value > 0:
        return (client_value / limit_value) * 100.0
    if (not direct) and limit_value > 0 and client_value > 0:
        return (limit_value / client_value) * 100.0

    denom = abs(limit_value) if limit_value != 0 else max(abs(client_value), 1e-9)
    if direct:
        return ((client_value - limit_value) / denom) * 100.0 + 100.0
    return ((limit_value - client_value) / denom) * 100.0 + 100.0


def zone_from_pi(pi: float | None, *, red_lt: float, yellow_hi: float) -> Zone:
    if pi is None:
        return "MISSING"
    if pi < red_lt:
        return "RED"
    if pi <= yellow_hi:
        return "YELLOW"
    return "GREEN"


def apply_activity_rules(
    *,
    critical_zones: list[Zone],
    supporting_zones: list[Zone],
    rules: AggregationRules = DEFAULT_RULES,
) -> Literal["RED", "YELLOW", "GREEN"]:
    # Treat missing Critical as YELLOW to force review.
    crit_red = any(z == "RED" for z in critical_zones)
    if crit_red:
        return "RED"

    supporting_red_count = sum(1 for z in supporting_zones if z == "RED")
    if supporting_red_count > rules.supporting_red_for_red:
        return "RED"

    crit_yellow_or_missing = any(z in {"YELLOW", "MISSING"} for z in critical_zones)
    if crit_yellow_or_missing:
        return "YELLOW"

    if supporting_red_count == rules.supporting_red_for_yellow:
        return "YELLOW"

    crit_all_green = all(z == "GREEN" for z in critical_zones) if critical_zones else False
    if crit_all_green and supporting_red_count < rules.supporting_red_for_yellow:
        return "GREEN"

    # Fallback for uncovered combinations (e.g., 3 supporting reds).
    return "YELLOW"

