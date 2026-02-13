"""Excel readers: threshold matrix and activity data."""

from pathlib import Path
from typing import Any

import pandas as pd

from ..utils import get_logger

_log = get_logger(__name__)


def read_threshold_matrix(
    path: str | Path,
    excel_config: dict[str, Any] | None = None,
) -> pd.DataFrame:
    """Read the threshold/classification matrix from Excel.

    Args:
        path: Path to the Excel file (e.g. Exercise Threshold Classifications.xlsx).
        excel_config: Optional dict with 'threshold_sheet', 'header_row'.

    Returns:
        DataFrame with threshold rows (columns depend on matrix layout).
    """
    config = excel_config or {}
    sheet = config.get("threshold_sheet", 0)
    header_row = config.get("header_row", 0)
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Threshold matrix not found: {path}")
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    _log.debug("Read threshold matrix: %s rows from %s", len(df), path)
    return df


def read_activity_data(
    path: str | Path,
    excel_config: dict[str, Any] | None = None,
) -> pd.DataFrame:
    """Read activity data from Excel.

    Args:
        path: Path to the Excel file (e.g. activity list or limits).
        excel_config: Optional dict with 'activity_sheet', 'header_row'.

    Returns:
        DataFrame with activity rows.
    """
    config = excel_config or {}
    sheet = config.get("activity_sheet", 0)
    header_row = config.get("header_row", 0)
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Activity data not found: {path}")
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    _log.debug("Read activity data: %s rows from %s", len(df), path)
    return df
