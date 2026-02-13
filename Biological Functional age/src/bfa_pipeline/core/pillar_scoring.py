"""Aggregate individual test scores into pillar functional ages."""

from __future__ import annotations
from typing import Optional, Dict
from dataclasses import dataclass

from .test_scoring import IndividualTestScores, TestResults
from .metabolic_scoring import calculate_metabolic_index


@dataclass
class PillarScores:
    """Functional ages for each of the 5 pillars."""
    vitality: Optional[float] = None
    strength: Optional[float] = None
    metabolic: Optional[float] = None
    mobility: Optional[float] = None
    cognitive: Optional[float] = None


def calculate_pillar_functional_ages(
    individual_scores: IndividualTestScores,
    tests: TestResults,
    weights: Dict[str, Dict[str, float]]
) -> PillarScores:
    """
    Calculate pillar functional ages from individual test scores.

    Args:
        individual_scores: Individual test functional ages and risk scores
        tests: Original test results (for age info)
        weights: Dict of pillar -> {test: weight} mappings

    Returns:
        PillarScores with functional ages for each pillar
    """
    pillars = PillarScores()

    # Vitality - already calculated as a combined functional age
    pillars.vitality = individual_scores.vo2_functional_age

    # Strength - weighted average of grip, STS, vertical jump
    strength_scores = []
    strength_weights = []

    if individual_scores.grip_functional_age is not None:
        strength_scores.append(individual_scores.grip_functional_age)
        strength_weights.append(weights['strength']['grip_strength'])

    if individual_scores.sts_functional_age is not None:
        strength_scores.append(individual_scores.sts_functional_age)
        strength_weights.append(weights['strength']['sts_power'])

    if individual_scores.vj_functional_age is not None:
        strength_scores.append(individual_scores.vj_functional_age)
        strength_weights.append(weights['strength']['vertical_jump'])

    if strength_scores:
        total_weight = sum(strength_weights)
        pillars.strength = sum(
            score * (weight / total_weight)
            for score, weight in zip(strength_scores, strength_weights)
        )

    # Mobility - weighted average of gait, TUG, SLS, sit&reach
    mobility_scores = []
    mobility_weights = []

    if individual_scores.gait_functional_age is not None:
        mobility_scores.append(individual_scores.gait_functional_age)
        mobility_weights.append(weights['mobility']['gait_speed'])

    if individual_scores.tug_functional_age is not None:
        mobility_scores.append(individual_scores.tug_functional_age)
        mobility_weights.append(weights['mobility']['tug'])

    if individual_scores.sls_functional_age is not None:
        mobility_scores.append(individual_scores.sls_functional_age)
        mobility_weights.append(weights['mobility']['single_leg_stance'])

    if individual_scores.sr_functional_age is not None:
        mobility_scores.append(individual_scores.sr_functional_age)
        mobility_weights.append(weights['mobility']['sit_and_reach'])

    if mobility_scores:
        total_weight = sum(mobility_weights)
        pillars.mobility = sum(
            score * (weight / total_weight)
            for score, weight in zip(mobility_scores, mobility_weights)
        )

    # Cognitive - weighted average of processing speed and working memory
    cognitive_scores = []
    cognitive_weights = []

    if individual_scores.processing_functional_age is not None:
        cognitive_scores.append(individual_scores.processing_functional_age)
        cognitive_weights.append(weights['cognitive']['processing_speed'])

    if individual_scores.memory_functional_age is not None:
        cognitive_scores.append(individual_scores.memory_functional_age)
        cognitive_weights.append(weights['cognitive']['working_memory'])

    if cognitive_scores:
        total_weight = sum(cognitive_weights)
        pillars.cognitive = sum(
            score * (weight / total_weight)
            for score, weight in zip(cognitive_scores, cognitive_weights)
        )

    # Metabolic - more complex
    # First calculate metabolic index from risk scores, then convert to functional age
    metabolic_scores_present = []
    metabolic_marker_scores = {}

    if individual_scores.apob_score is not None:
        metabolic_marker_scores['apob'] = individual_scores.apob_score
        metabolic_scores_present.append('apob')

    if individual_scores.homa_score is not None:
        metabolic_marker_scores['homa_ir'] = individual_scores.homa_score
        metabolic_scores_present.append('homa_ir')

    if individual_scores.hba1c_score is not None:
        metabolic_marker_scores['hba1c'] = individual_scores.hba1c_score
        metabolic_scores_present.append('hba1c')

    if individual_scores.whtr_score is not None:
        metabolic_marker_scores['whtr'] = individual_scores.whtr_score
        metabolic_scores_present.append('whtr')

    if individual_scores.hscrp_score is not None:
        metabolic_marker_scores['hscrp'] = individual_scores.hscrp_score
        metabolic_scores_present.append('hscrp')

    if individual_scores.body_fat_score is not None:
        metabolic_marker_scores['body_fat_pct'] = individual_scores.body_fat_score
        metabolic_scores_present.append('body_fat_pct')

    if metabolic_scores_present:
        # Calculate metabolic index (0-100 weighted score)
        metabolic_index = calculate_metabolic_index(
            apob_score=metabolic_marker_scores.get('apob', 5.0),
            homa_ir_score=metabolic_marker_scores.get('homa_ir', 5.0),
            hba1c_score=metabolic_marker_scores.get('hba1c', 5.0),
            whtr_score=metabolic_marker_scores.get('whtr', 5.0),
            hscrp_score=metabolic_marker_scores.get('hscrp', 5.0),
            body_fat_score=metabolic_marker_scores.get('body_fat_pct', 5.0),
            weights=weights['metabolic']
        )

        # Convert metabolic index to functional age
        # Based on Excel formula: Age + delta
        # where delta = (index - baseline) * (span / baseline)
        # For simplicity, using a linear relationship:
        # - Index of 5 (optimal) = chronological age
        # - Index of 100 (worst) = chronological age + span
        # Using span of 30 years as a reasonable range

        age = tests.age
        baseline = 5.0  # Optimal index
        span = 30.0  # Years of aging represented by index range

        # Calculate delta from optimal
        delta = (metabolic_index - baseline) * (span / baseline)

        # Clamp delta to reasonable bounds
        delta = max(-span, min(span, delta))

        pillars.metabolic = age + delta

    return pillars


def check_completeness(pillars: PillarScores) -> bool:
    """
    Check if all required pillars are present.

    Args:
        pillars: Calculated pillar scores

    Returns:
        True if all 5 pillars are present, False otherwise
    """
    return all([
        pillars.vitality is not None,
        pillars.strength is not None,
        pillars.metabolic is not None,
        pillars.mobility is not None,
        pillars.cognitive is not None
    ])
