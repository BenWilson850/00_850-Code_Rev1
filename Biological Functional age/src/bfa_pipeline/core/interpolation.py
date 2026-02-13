"""Interpolation utilities for normative data lookups."""

from __future__ import annotations
from typing import List, Tuple
import numpy as np


def linear_interpolate(x: float, x_values: List[float], y_values: List[float]) -> float:
    """
    Perform linear interpolation to find y at position x.

    Args:
        x: The value to interpolate for
        x_values: List of x coordinates (must be sorted)
        y_values: List of y coordinates corresponding to x_values

    Returns:
        Interpolated y value at position x

    Example:
        >>> linear_interpolate(42, [40, 45], [49, 47])
        48.2  # Interpolated grip strength for age 42
    """
    if len(x_values) != len(y_values):
        raise ValueError("x_values and y_values must have same length")

    if len(x_values) < 2:
        raise ValueError("Need at least 2 points for interpolation")

    # If x is outside the range, use nearest value (extrapolation)
    if x <= x_values[0]:
        return y_values[0]
    if x >= x_values[-1]:
        return y_values[-1]

    # Use numpy for interpolation
    return float(np.interp(x, x_values, y_values))


def interpolate_from_age_ranges(
    age: float,
    age_ranges: List[Tuple[float, float]],
    values: List[float]
) -> float:
    """
    Interpolate between range midpoints for decade-based normative data.

    Args:
        age: Client's age
        age_ranges: List of (min_age, max_age) tuples, e.g., [(20, 29), (30, 39)]
        values: List of normative values for each range

    Returns:
        Interpolated value for the given age

    Example:
        >>> interpolate_from_age_ranges(42, [(20, 29), (30, 39), (40, 49)], [45.4, 44.0, 42.4])
        42.9  # VO2 max for 42-year-old
    """
    if len(age_ranges) != len(values):
        raise ValueError("age_ranges and values must have same length")

    # Calculate midpoints of ranges
    midpoints = [(min_age + max_age) / 2 for min_age, max_age in age_ranges]

    # Use linear interpolation between midpoints
    return linear_interpolate(age, midpoints, values)


def find_functional_age(
    test_value: float,
    age_values: List[float],
    norm_values: List[float],
    reverse: bool = False
) -> float:
    """
    Find the functional age where a test result would be normal.

    This reverses the normal interpolation: given a test result, find the age
    where that result would be average.

    Args:
        test_value: Client's test result
        age_values: List of ages in normative data
        norm_values: List of normative values for each age
        reverse: If True, higher test values = older (e.g., TUG time)
                 If False, higher test values = younger (e.g., grip strength)

    Returns:
        Functional age (age at which this test result would be normal)

    Example:
        >>> find_functional_age(44, [20, 25, 30, 35, 40], [48, 50, 51, 50, 49], False)
        37.5  # 44kg grip strength is normal for ~37.5 year old
    """
    if len(age_values) != len(norm_values):
        raise ValueError("age_values and norm_values must have same length")

    if len(age_values) < 2:
        raise ValueError("Need at least 2 points for interpolation")

    # For reverse metrics (higher = worse, like TUG), swap the interpolation
    if reverse:
        # Interpolate age from the test value
        # Clamp to range bounds
        if test_value <= min(norm_values):
            return min(age_values)
        if test_value >= max(norm_values):
            return max(age_values)
        return float(np.interp(test_value, norm_values, age_values))
    else:
        # For normal metrics (higher = better, like grip strength)
        # Clamp to range bounds
        if test_value >= max(norm_values):
            return min(age_values)
        if test_value <= min(norm_values):
            return max(age_values)
        # Reverse the norm_values since we're going backwards
        return float(np.interp(test_value, list(reversed(norm_values)), list(reversed(age_values))))


def clamp(value: float, min_val: float, max_val: float) -> float:
    """
    Clamp a value between min and max.

    Args:
        value: Value to clamp
        min_val: Minimum allowed value
        max_val: Maximum allowed value

    Returns:
        Clamped value
    """
    return max(min_val, min(max_val, value))
