"""Individual test scoring to functional ages."""

from __future__ import annotations
from typing import Optional, Dict
from dataclasses import dataclass

from .interpolation import (
    linear_interpolate,
    interpolate_from_age_ranges,
    find_functional_age,
    clamp
)
from .metabolic_scoring import (
    score_metabolic_marker,
    score_hscrp_special,
    calculate_metabolic_index
)
from ..io.normative_reader import NormativeData


@dataclass
class TestResults:
    """Container for all client test results."""
    # Client info
    name: str
    age: float
    gender: str  # 'Male' or 'Female'

    # Vitality tests
    vo2_max: Optional[float] = None
    fev1: Optional[float] = None

    # Strength tests
    grip_strength: Optional[float] = None
    sts_power: Optional[float] = None
    vertical_jump: Optional[float] = None

    # Metabolic tests
    body_fat_pct: Optional[float] = None
    whtr: Optional[float] = None
    fasting_glucose: Optional[float] = None
    hba1c: Optional[float] = None
    homa_ir: Optional[float] = None
    apob: Optional[float] = None
    hscrp: Optional[float] = None

    # Mobility tests
    gait_speed: Optional[float] = None
    tug: Optional[float] = None
    single_leg_stance: Optional[float] = None
    sit_and_reach: Optional[float] = None

    # Cognitive tests (pre-normalized as SD from norm)
    processing_speed: Optional[float] = None
    working_memory: Optional[float] = None


@dataclass
class IndividualTestScores:
    """Functional ages or scores for individual tests."""
    # Physical test functional ages
    vo2_functional_age: Optional[float] = None
    fev1_functional_age: Optional[float] = None
    grip_functional_age: Optional[float] = None
    sts_functional_age: Optional[float] = None
    vj_functional_age: Optional[float] = None
    gait_functional_age: Optional[float] = None
    tug_functional_age: Optional[float] = None
    sls_functional_age: Optional[float] = None
    sr_functional_age: Optional[float] = None
    processing_functional_age: Optional[float] = None
    memory_functional_age: Optional[float] = None

    # Metabolic risk scores (0-100)
    whtr_score: Optional[float] = None
    hba1c_score: Optional[float] = None
    homa_score: Optional[float] = None
    apob_score: Optional[float] = None
    hscrp_score: Optional[float] = None
    body_fat_score: Optional[float] = None


def score_physical_test(
    test_value: float,
    age: float,
    gender: str,
    norm_data: Dict[str, list],
    reverse: bool = False
) -> float:
    """
    Score a physical test by finding functional age via interpolation.

    Args:
        test_value: Client's test result
        age: Client's chronological age
        gender: 'Male' or 'Female'
        norm_data: Dict with 'Male' and 'Female' keys containing (age, value) tuples
        reverse: True if higher values = older (e.g., TUG time)

    Returns:
        Functional age for this test
    """
    if gender not in norm_data:
        raise ValueError(f"Gender '{gender}' not found in normative data")

    gender_norms = norm_data[gender]

    # Check if data is in (age, value) or ((age_min, age_max), value) format
    if isinstance(gender_norms[0][0], tuple):
        # Age range format - convert to ages and values for interpolation
        ages = [(min_age + max_age) / 2 for (min_age, max_age), _ in gender_norms]
        values = [val for _, val in gender_norms]
    else:
        # Point format
        ages = [age for age, _ in gender_norms]
        values = [val for _, val in gender_norms]

    return find_functional_age(test_value, ages, values, reverse=reverse)


def score_vo2_max_vitality_component(
    vo2_value: float,
    age: float,
    gender: str,
    norm_data: NormativeData
) -> float:
    """
    Score VO2 max component for vitality using the special formula.

    This uses the formula from Excel that calculates a ratio to optimal VO2.

    Args:
        vo2_value: Client's VO2 max
        age: Client's age
        gender: Client's gender

    Returns:
        VO2 performance ratio (0-1 scale, where 1 = optimal)
    """
    # Get age/gender-specific normative VO2
    gender_norms = norm_data.vo2_max[gender]
    age_ranges = [age_range for age_range, _ in gender_norms]
    vo2_norms = [vo2 for _, vo2 in gender_norms]

    # Interpolate to get expected VO2 for this age/gender
    vo2_norm = interpolate_from_age_ranges(age, age_ranges, vo2_norms)

    # Calculate performance ratio (client / norm)
    return vo2_value / vo2_norm


def calculate_vitality_functional_age(
    vo2_value: float,
    fev1_value: float,
    age: float,
    gender: str,
    norm_data: NormativeData,
    vo2_weight: float = 0.7,
    fev1_weight: float = 0.3
) -> float:
    """
    Calculate vitality functional age using the special formula.

    Formula: Age + ((1 - ((VO2/VO2_norm) * 0.7) - ((FEV1/100) * 0.3)) * 100)

    Args:
        vo2_value: Client's VO2 max (ml/kg/min)
        fev1_value: Client's FEV1 (% predicted)
        age: Client's chronological age
        gender: Client's gender
        norm_data: Normative data
        vo2_weight: Weight for VO2 component (default 0.7)
        fev1_weight: Weight for FEV1 component (default 0.3)

    Returns:
        Vitality functional age
    """
    # Get VO2 performance ratio
    vo2_ratio = score_vo2_max_vitality_component(vo2_value, age, gender, norm_data)

    # FEV1 is already % predicted, so 100% is optimal
    fev1_ratio = fev1_value / 100.0

    # Calculate weighted performance
    weighted_performance = (vo2_ratio * vo2_weight) + (fev1_ratio * fev1_weight)

    # Calculate deficit from optimal
    deficit = 1.0 - weighted_performance

    # Convert to functional age
    functional_age = age + (deficit * 100.0)

    return functional_age


def score_cognitive_test(
    sd_value: float,
    age: float
) -> float:
    """
    Convert cognitive test SD to functional age.

    Cognitive tests are pre-normalized as standard deviations from the norm.
    Negative SD = worse than average = older functional age.

    Formula from Excel: MAX(18, Age - (SD * 25))

    Args:
        sd_value: Standard deviations from norm (can be negative)
        age: Chronological age

    Returns:
        Cognitive functional age
    """
    # Each SD = 25 years of aging
    functional_age = age - (sd_value * 25.0)

    # Minimum functional age is 18
    return max(18.0, functional_age)


def score_all_tests(
    tests: TestResults,
    norm_data: NormativeData,
    metabolic_weights: Dict[str, float]
) -> IndividualTestScores:
    """
    Score all individual tests for a client.

    Args:
        tests: Client's test results
        norm_data: Normative reference data
        metabolic_weights: Weights for metabolic sub-tests

    Returns:
        IndividualTestScores with functional ages and risk scores
    """
    scores = IndividualTestScores()

    # Vitality - special handling
    if tests.vo2_max is not None and tests.fev1 is not None:
        scores.vo2_functional_age = calculate_vitality_functional_age(
            tests.vo2_max,
            tests.fev1,
            tests.age,
            tests.gender,
            norm_data
        )

    # Strength tests
    if tests.grip_strength is not None:
        scores.grip_functional_age = score_physical_test(
            tests.grip_strength,
            tests.age,
            tests.gender,
            norm_data.grip_strength,
            reverse=False
        )

    if tests.sts_power is not None:
        scores.sts_functional_age = score_physical_test(
            tests.sts_power,
            tests.age,
            tests.gender,
            norm_data.sts_power,
            reverse=False
        )

    if tests.vertical_jump is not None:
        scores.vj_functional_age = score_physical_test(
            tests.vertical_jump,
            tests.age,
            tests.gender,
            norm_data.vertical_jump,
            reverse=False
        )

    # Mobility tests
    if tests.gait_speed is not None:
        scores.gait_functional_age = score_physical_test(
            tests.gait_speed,
            tests.age,
            tests.gender,
            norm_data.gait_speed,
            reverse=False
        )

    if tests.tug is not None:
        scores.tug_functional_age = score_physical_test(
            tests.tug,
            tests.age,
            tests.gender,
            norm_data.tug,
            reverse=True  # Higher TUG time = worse
        )

    if tests.single_leg_stance is not None:
        # SLS has no gender split
        sls_data = {'Male': norm_data.single_leg_stance, 'Female': norm_data.single_leg_stance}
        scores.sls_functional_age = score_physical_test(
            tests.single_leg_stance,
            tests.age,
            tests.gender,
            sls_data,
            reverse=False
        )

    if tests.sit_and_reach is not None:
        scores.sr_functional_age = score_physical_test(
            tests.sit_and_reach,
            tests.age,
            tests.gender,
            norm_data.sit_and_reach,
            reverse=False
        )

    # Cognitive tests (pre-normalized)
    if tests.processing_speed is not None:
        scores.processing_functional_age = score_cognitive_test(
            tests.processing_speed,
            tests.age
        )

    if tests.working_memory is not None:
        scores.memory_functional_age = score_cognitive_test(
            tests.working_memory,
            tests.age
        )

    # Metabolic markers - risk scoring (0-100)
    if tests.whtr is not None:
        ranges = norm_data.metabolic_ranges.get('WHtR', {})
        scores.whtr_score = score_metabolic_marker(
            tests.whtr,
            ranges.get('Low risk', '<0.5'),
            ranges.get('Normal', '0.5-0.59'),
            ranges.get('Elevated', '>0.6')
        )

    if tests.hba1c is not None:
        ranges = norm_data.metabolic_ranges.get('HbA1c', {})
        scores.hba1c_score = score_metabolic_marker(
            tests.hba1c,
            ranges.get('Low risk', '4.0-5.4'),
            ranges.get('Normal', '5.5-5.6'),
            ranges.get('Elevated', '>5.7')
        )

    if tests.homa_ir is not None:
        ranges = norm_data.metabolic_ranges.get('Homa', {})
        scores.homa_score = score_metabolic_marker(
            tests.homa_ir,
            ranges.get('Low risk', '<1.5'),
            ranges.get('Normal', '1.5-2.4'),
            ranges.get('Elevated', '>2.5')
        )

    if tests.apob is not None:
        ranges = norm_data.metabolic_ranges.get('ApoB', {})
        scores.apob_score = score_metabolic_marker(
            tests.apob,
            ranges.get('Low risk', '<90'),
            ranges.get('Normal', '90-129'),
            ranges.get('Elevated', '>130')
        )

    if tests.hscrp is not None:
        # hsCRP uses special scoring
        scores.hscrp_score = score_hscrp_special(tests.hscrp)

    if tests.body_fat_pct is not None:
        # Body fat uses age/gender-specific ranges
        gender_ranges = norm_data.body_fat_ranges[tests.gender]
        # Find appropriate age range
        for (age_min, age_max), (healthy_max, overweight_max, obese_min) in gender_ranges:
            if age_min <= tests.age <= age_max:
                # Create pseudo risk ranges
                low_risk = f'<{healthy_max}'
                normal = f'{healthy_max}-{overweight_max}'
                elevated = f'>{obese_min}'
                scores.body_fat_score = score_metabolic_marker(
                    tests.body_fat_pct,
                    low_risk,
                    normal,
                    elevated
                )
                break

    return scores
