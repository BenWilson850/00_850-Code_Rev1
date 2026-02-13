"""Write BFA calculation results to CSV/Excel."""

from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd

from ..core.test_scoring import TestResults
from ..core.pillar_scoring import PillarScores


def format_value(value: Any) -> Any:
    """
    Format a value for output, handling None and rounding floats.

    Args:
        value: Value to format

    Returns:
        Formatted value
    """
    if value is None:
        return "INCOMPLETE"
    if isinstance(value, float):
        return round(value, 2)
    return value


def create_output_row(
    tests: TestResults,
    pillars: PillarScores,
    bfa: float | None,
    healthspan_index: float | None,
    healthspan_category: str | None
) -> Dict[str, Any]:
    """
    Create a single output row with all data.

    Args:
        tests: Client test results
        pillars: Calculated pillar scores
        bfa: Biological Functional Age
        healthspan_index: Healthspan Index score
        healthspan_category: Healthspan category name

    Returns:
        Dict with all output columns
    """
    return {
        # Client metadata
        'Name': tests.name,
        'Age': tests.age,
        'Gender': tests.gender,

        # Raw test results (16 tests)
        'VO2_Max': tests.vo2_max,
        'FEV1': tests.fev1,
        'Grip_Strength': tests.grip_strength,
        'STS_Power': tests.sts_power,
        'Vertical_Jump': tests.vertical_jump,
        'Body_Fat_Pct': tests.body_fat_pct,
        'WHtR': tests.whtr,
        'Fasting_Glucose': tests.fasting_glucose,
        'HbA1c': tests.hba1c,
        'HOMA_IR': tests.homa_ir,
        'ApoB': tests.apob,
        'hsCRP': tests.hscrp,
        'Gait_Speed': tests.gait_speed,
        'TUG': tests.tug,
        'Single_Leg_Stance': tests.single_leg_stance,
        'Sit_And_Reach': tests.sit_and_reach,
        'Processing_Speed': tests.processing_speed,
        'Working_Memory': tests.working_memory,

        # Pillar functional ages (5 pillars)
        'Vitality_Functional_Age': format_value(pillars.vitality),
        'Strength_Functional_Age': format_value(pillars.strength),
        'Metabolic_Functional_Age': format_value(pillars.metabolic),
        'Mobility_Functional_Age': format_value(pillars.mobility),
        'Cognitive_Functional_Age': format_value(pillars.cognitive),

        # Final outputs
        'Biological_Functional_Age': format_value(bfa),
        'Healthspan_Index': format_value(healthspan_index),
        'Healthspan_Category': healthspan_category if healthspan_category else "INCOMPLETE"
    }


def write_results_csv(
    results: List[Dict[str, Any]],
    output_path: str | Path
) -> None:
    """
    Write results to CSV file.

    Args:
        results: List of result dicts (one per client)
        output_path: Path to output CSV file
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(results)
    df.to_csv(output_path, index=False)

    print(f"Results written to: {output_path}")
    print(f"Total clients processed: {len(results)}")


def write_results_excel(
    results: List[Dict[str, Any]],
    output_path: str | Path
) -> None:
    """
    Write results to Excel file.

    Args:
        results: List of result dicts (one per client)
        output_path: Path to output Excel file
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"Results written to: {output_path}")
    print(f"Total clients processed: {len(results)}")
