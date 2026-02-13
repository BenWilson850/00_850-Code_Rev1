"""Pipeline core: engine, assessment, aggregation."""

from .engine import run_pipeline
from .assess import assess_activities
from .aggregation import aggregate_results

__all__ = ["run_pipeline", "assess_activities", "aggregate_results"]
