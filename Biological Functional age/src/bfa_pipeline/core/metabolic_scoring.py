"""Metabolic marker risk-based scoring."""

from __future__ import annotations
from typing import Dict, Tuple
import re


def parse_risk_range(range_str: str) -> Tuple[float, float]:
    """
    Parse risk range string like '<90', '90-129', or '>130' into (min, max).

    Args:
        range_str: String like '<90', '90 - 129', '>130', '4.0 - 5.4'

    Returns:
        Tuple of (min_value, max_value)
        Uses -inf for open lower bounds and +inf for open upper bounds
    """
    range_str = range_str.strip().replace('–', '-').replace(' ', '')

    if range_str.startswith('<'):
        # '<90' means -inf to 90
        max_val = float(re.sub(r'[^\d.]', '', range_str))
        return (float('-inf'), max_val)
    elif range_str.startswith('>'):
        # '>130' means 130 to +inf
        min_val = float(re.sub(r'[^\d.]', '', range_str))
        return (min_val, float('inf'))
    elif '-' in range_str:
        # '90-129' means 90 to 129
        parts = range_str.split('-')
        return (float(parts[0]), float(parts[1]))
    else:
        # Single value - treat as point
        val = float(range_str)
        return (val, val)


def score_metabolic_marker(
    value: float,
    low_risk_range: str,
    normal_range: str,
    elevated_range: str
) -> float:
    """
    Score a metabolic marker on 0-100 scale based on risk categories.

    Scoring logic (matching Excel LET formula):
    - Low risk range: score ~5 (optimal)
    - Normal range: score 5-35 (warning zone)
    - Elevated range: score 35-100 (high risk)

    Args:
        value: Client's test result
        low_risk_range: Low risk range string (e.g., '<90')
        normal_range: Normal range string (e.g., '90-129')
        elevated_range: Elevated range string (e.g., '>130')

    Returns:
        Score from 0-100 (higher = worse)
    """
    # Parse ranges
    low_min, low_max = parse_risk_range(low_risk_range)
    norm_min, norm_max = parse_risk_range(normal_range)
    elev_min, elev_max = parse_risk_range(elevated_range)

    # Calculate span for scaling (avoid division by zero)
    span = max(norm_max - low_max, 0.000001)

    # Score based on which range the value falls into
    if value <= low_max:
        # In low risk range: score = 5
        return 5.0

    elif value < elev_min:
        # In normal range: score = 5 to 35
        # Linear interpolation from 5 at low_max to 35 at elev_min
        numerator = value - low_max
        denominator = max(elev_min - low_max, 0.000001)
        score = 5.0 + 30.0 * (numerator / denominator)
        return score

    else:
        # In elevated range: score = 35 to 100
        # Linear interpolation from 35 at elev_min upwards
        numerator = value - elev_min
        denominator = 2 * span
        score = 35.0 + 65.0 * (numerator / denominator)
        # Clamp at 100
        return min(100.0, max(0.0, score))


def calculate_metabolic_index(
    apob_score: float,
    homa_ir_score: float,
    hba1c_score: float,
    whtr_score: float,
    hscrp_score: float,
    body_fat_score: float,
    weights: Dict[str, float]
) -> float:
    """
    Calculate weighted metabolic index from individual marker scores.

    Args:
        apob_score: ApoB risk score (0-100)
        homa_ir_score: HOMA-IR risk score (0-100)
        hba1c_score: HbA1c risk score (0-100)
        whtr_score: WHtR risk score (0-100)
        hscrp_score: hsCRP risk score (0-100)
        body_fat_score: Body fat % risk score (0-100)
        weights: Dict with keys: apob, homa_ir, hba1c, whtr, hscrp, body_fat_pct

    Returns:
        Weighted metabolic index (0-100 scale)
    """
    index = (
        (weights['apob'] / 100.0) * apob_score +
        (weights['homa_ir'] / 100.0) * homa_ir_score +
        (weights['hba1c'] / 100.0) * hba1c_score +
        (weights['whtr'] / 100.0) * whtr_score +
        (weights['hscrp'] / 100.0) * hscrp_score +
        (weights['body_fat_pct'] / 100.0) * body_fat_score
    )
    return index


def score_hscrp_special(hscrp_value: float) -> float:
    """
    Special scoring for hsCRP based on Excel formula.

    Formula from Excel:
    IF(C13<1, 10*(C13/1),
       IF(C13<3, 10+40*(C13-1)/2,
          50+50*(MIN(C13,10)-3)/7))

    Args:
        hscrp_value: hsCRP value in mg/L

    Returns:
        Score from 0-100
    """
    if hscrp_value < 1:
        return 10.0 * (hscrp_value / 1.0)
    elif hscrp_value < 3:
        return 10.0 + 40.0 * ((hscrp_value - 1.0) / 2.0)
    else:
        clamped = min(hscrp_value, 10.0)
        return 50.0 + 50.0 * ((clamped - 3.0) / 7.0)
