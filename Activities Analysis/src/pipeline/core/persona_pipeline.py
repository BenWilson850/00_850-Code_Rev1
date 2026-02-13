"""Persona workbook pipeline: assess 5y/10y predictions against activity limits."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

from ..io import read_persona_workbook, read_docx_spec
from ..report import write_persona_report
from ..utils import get_logger
from .persona_logic import (
    DEFAULT_RULES,
    AggregationRules,
    apply_activity_rules,
    compute_pi,
    infer_importance,
    parse_limit,
    zone_from_pi,
)
from .persona_match import infer_test_id, is_meta_column


_log = get_logger(__name__)


def _load_matrix(path: str | Path, *, sheet: int | str = 0, header_row: int = 0) -> pd.DataFrame:
    df = pd.read_excel(Path(path), sheet_name=sheet, header=header_row)
    if "Activity" not in df.columns:
        raise ValueError(f"Expected an 'Activity' column in {path}")
    # Normalize Activity column to string for stable indexing.
    df["Activity"] = df["Activity"].astype(str).str.strip()
    df = df.set_index("Activity")
    return df


def run_persona_pipeline(
    config: dict[str, Any],
    *,
    clients_path: str | Path,
    limits_path: str | Path,
    classifications_path: str | Path,
    docx_path: str | Path | None = None,
    output_dir: str | Path = "out",
) -> Path:
    excel_cfg = config.get("excel", {})
    client_cfg = config.get("client", {})
    status_cfg = config.get("status", {})
    output_cfg = config.get("output", {})

    sheet = excel_cfg.get("matrix_sheet", 0)
    header_row = int(excel_cfg.get("header_row", 0))

    _log.info("Reading limits matrix from %s", limits_path)
    limits_df = _load_matrix(limits_path, sheet=sheet, header_row=header_row)
    _log.info("Reading classifications matrix from %s", classifications_path)
    class_df = _load_matrix(classifications_path, sheet=sheet, header_row=header_row)

    # Determine test columns from the limits matrix; assume same columns in classifications.
    test_cols = [c for c in limits_df.columns if not is_meta_column(c)]
    if not test_cols:
        raise ValueError("No test columns detected in limits matrix")

    rules = _rules_from_docx(docx_path) if docx_path else DEFAULT_RULES

    _log.info("Reading client workbook from %s", clients_path)
    clients = read_persona_workbook(clients_path, client_cfg)

    red_lt = float(status_cfg.get("red_lt", 90.0))
    yellow_hi = float(status_cfg.get("yellow_hi", 110.0))

    client_reports: list[dict[str, Any]] = []
    for client_sheet in clients:
        client = client_sheet.client
        rows_by_activity: dict[str, dict[str, Any]] = {}

        for activity in limits_df.index.tolist():
            for horizon in ("5", "10"):
                critical_zones = []
                supporting_zones = []
                critical_failures: list[str] = []
                supporting_failures: list[str] = []

                for col in test_cols:
                    raw_limit = limits_df.at[activity, col]
                    if raw_limit is None or (isinstance(raw_limit, float) and pd.isna(raw_limit)):
                        continue
                    raw_importance = None
                    if activity in class_df.index and col in class_df.columns:
                        raw_importance = class_df.at[activity, col]
                    importance = infer_importance(raw_importance) or "Supporting"

                    test_id = infer_test_id(col)
                    client_test = client_sheet.tests.get(test_id)
                    client_value = None if not client_test else client_test.get(horizon)

                    limit = parse_limit(raw_limit, gender=client.gender)
                    if not limit:
                        continue
                    pi = compute_pi(client_value, limit)
                    zone = zone_from_pi(pi, red_lt=red_lt, yellow_hi=yellow_hi)

                    display_name = str(col).replace("\n", " ").strip()
                    if importance == "Critical":
                        critical_zones.append(zone)
                        if zone != "GREEN":
                            critical_failures.append(f"{display_name} ({zone})")
                    else:
                        supporting_zones.append(zone)
                        if zone == "RED":
                            supporting_failures.append(f"{display_name} ({zone})")

                final_status = apply_activity_rules(
                    critical_zones=critical_zones,
                    supporting_zones=supporting_zones,
                    rules=rules,
                )

                # Record output row. One row per activity with 5y/10y columns; we fill in after loop.
                row = rows_by_activity.get(activity)
                if horizon == "5":
                    row = {
                        "Activity": activity,
                        "5 Year Critical Failures": ", ".join(critical_failures),
                        "5 Year Supporting Failures": ", ".join(supporting_failures),
                        "5 Year Final Status": final_status,
                        "10 Year Critical Failures": "",
                        "10 Year Supporting Failures": "",
                        "10 Year Final Status": "",
                    }
                    rows_by_activity[activity] = row
                else:
                    if row is None:
                        row = {
                            "Activity": activity,
                            "5 Year Critical Failures": "",
                            "5 Year Supporting Failures": "",
                            "5 Year Final Status": "",
                            "10 Year Critical Failures": "",
                            "10 Year Supporting Failures": "",
                            "10 Year Final Status": "",
                        }
                        rows_by_activity[activity] = row
                    row["10 Year Critical Failures"] = ", ".join(critical_failures)
                    row["10 Year Supporting Failures"] = ", ".join(supporting_failures)
                    row["10 Year Final Status"] = final_status

        rows = [rows_by_activity[a] for a in limits_df.index.tolist() if a in rows_by_activity]
        client_reports.append({"client": client, "rows": rows})

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_name = config.get("paths", {}).get("report_name", "persona_activity_report.xlsx")
    report_path = output_dir / report_name

    spec = read_docx_spec(docx_path) if docx_path else None
    write_persona_report(
        client_reports,
        report_path,
        sheet_name_template=output_cfg.get("sheet_name_template", "{name}_{age}"),
        max_sheet_name_len=int(output_cfg.get("max_sheet_name_len", 31)),
        spec=spec,
    )
    return report_path


def _rules_from_docx(docx_path: str | Path | None) -> AggregationRules:
    if not docx_path:
        return DEFAULT_RULES
    spec = read_docx_spec(docx_path)
    if not spec or not spec.get("tables"):
        return DEFAULT_RULES

    # Table 0 is the rule table in the sample docx.
    table0 = spec["tables"][0] if spec["tables"] else []
    supporting_red_for_red = DEFAULT_RULES.supporting_red_for_red
    supporting_red_for_yellow = DEFAULT_RULES.supporting_red_for_yellow

    for row in table0[1:]:
        if len(row) < 2:
            continue
        trigger = str(row[1]).lower()
        if "supporting" in trigger and ">" in trigger and "red" in trigger:
            # e.g. ">3 Supporting are RED (<90%)"
            nums = [int(n) for n in re.findall(r"(\d+)", trigger)]
            if nums:
                supporting_red_for_red = max(nums)
        if "supporting" in trigger and " are" in trigger and "red" in trigger and ">" not in trigger:
            # e.g. "2 Supporting are RED (<90)"
            nums = [int(n) for n in re.findall(r"(\d+)", trigger)]
            if nums:
                supporting_red_for_yellow = max(nums)

    return AggregationRules(
        supporting_red_for_red=supporting_red_for_red,
        supporting_red_for_yellow=supporting_red_for_yellow,
    )
