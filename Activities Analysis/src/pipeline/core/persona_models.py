"""Models and helpers for the persona workbook matrix pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Literal


Importance = Literal["Critical", "Supporting"]
Zone = Literal["RED", "YELLOW", "GREEN", "MISSING"]


@dataclass(frozen=True)
class LimitSpec:
    op: Literal[">", ">=", "<", "<="]
    value: float


@dataclass(frozen=True)
class TestAssessment:
    test_name: str
    test_id: str
    importance: Importance
    client_value: float | None
    limit: LimitSpec
    pi: float | None
    zone: Zone


@dataclass(frozen=True)
class ActivityAssessment:
    activity: str
    horizon: Literal["5", "10"]
    critical_tests: list[TestAssessment]
    supporting_tests: list[TestAssessment]
    final_status: Literal["RED", "YELLOW", "GREEN"]
    critical_failures: list[str]
    supporting_failures: list[str]


@dataclass(frozen=True)
class ClientInfo:
    name: str
    age: int | None
    gender: Literal["M", "F"] | None
    sheet_name: str


@dataclass(frozen=True)
class ClientReport:
    client: ClientInfo
    rows: list[dict[str, Any]]

