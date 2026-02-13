"""Calculate Healthspan Index from BFA."""

from __future__ import annotations
from typing import Dict, Tuple, Optional

from .interpolation import clamp


def calculate_healthspan_index(
    chronological_age: float,
    biological_age: Optional[float],
    config: Dict
) -> Tuple[Optional[float], Optional[str]]:
    """
    Calculate Healthspan Index (300-850) and category.

    Formula: 670 + (6.5 * (Chronological_Age - BFA))
    Where BFA is clamped between 25-90

    Args:
        chronological_age: Client's actual age
        biological_age: Calculated BFA (or None if incomplete)
        config: Config dict with base_score, points_per_year, min_score, max_score, bfa_clamp_min, bfa_clamp_max

    Returns:
        Tuple of (healthspan_index, category) or (None, None) if BFA is None
    """
    if biological_age is None:
        return None, None

    # Clamp BFA for calculation
    bfa_clamped = clamp(
        biological_age,
        config['bfa_clamp_min'],
        config['bfa_clamp_max']
    )

    # Calculate age difference
    age_diff = chronological_age - bfa_clamped

    # Calculate index
    index = config['base_score'] + (config['points_per_year'] * age_diff)

    # Clamp to range
    index = clamp(index, config['min_score'], config['max_score'])

    return index, None  # Category will be determined separately


def categorize_healthspan_index(
    index: Optional[float],
    categories: Dict[str, Dict]
) -> Optional[str]:
    """
    Determine Healthspan Index category.

    Args:
        index: Healthspan Index score
        categories: Dict of category_name -> {min, max, description}

    Returns:
        Category name (Critical/Poor/Fair/Average/Good/Excellent/Elite) or None
    """
    if index is None:
        return None

    for category_name, info in categories.items():
        if info['min'] <= index <= info['max']:
            return category_name

    return "Unknown"
