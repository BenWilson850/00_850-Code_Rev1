"""Calculate Biological Functional Age from pillar scores."""

from __future__ import annotations
from typing import Dict, Optional

from .pillar_scoring import PillarScores, check_completeness


def calculate_bfa(
    pillars: PillarScores,
    pillar_weights: Dict[str, float]
) -> Optional[float]:
    """
    Calculate Biological Functional Age from pillar functional ages.

    Args:
        pillars: Pillar functional ages
        pillar_weights: Dict with pillar weights (metabolic, vitality, strength, mobility, cognitive)

    Returns:
        BFA as weighted average of pillar ages, or None if incomplete
    """
    # Check if all pillars are present
    if not check_completeness(pillars):
        return None

    # Calculate weighted average
    # Weights should sum to 100
    total_weight = sum(pillar_weights.values())

    bfa = (
        (pillars.metabolic * pillar_weights['metabolic'] / total_weight) +
        (pillars.vitality * pillar_weights['vitality'] / total_weight) +
        (pillars.strength * pillar_weights['strength'] / total_weight) +
        (pillars.mobility * pillar_weights['mobility'] / total_weight) +
        (pillars.cognitive * pillar_weights['cognitive'] / total_weight)
    )

    return bfa
