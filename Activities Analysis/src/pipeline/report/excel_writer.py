"""Excel report writer: write aggregated results to Excel."""

from pathlib import Path
from typing import Any

import pandas as pd

from ..utils import get_logger

_log = get_logger(__name__)


def write_report(
    aggregated: dict[str, Any],
    path: str | Path,
    spec: dict[str, Any] | None = None,
) -> None:
    """Write the pipeline report to an Excel file.

    Creates sheets for by_activity, by_threshold summary, and optional
    spec summary if a DOCX spec was provided.

    Args:
        aggregated: Output from aggregate_results.
        path: Output file path (e.g. out/matrix_report.xlsx).
        spec: Optional DOCX spec content for a narrative sheet.
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # By-activity sheet: flatten one row per activity with outcome summary
        by_activity = aggregated.get("by_activity", [])
        if by_activity:
            rows = []
            for row in by_activity:
                outcomes = row.get("threshold_outcomes", [])
                met = [o.get("threshold_id") for o in outcomes if o.get("met")]
                rows.append({**{k: v for k, v in row.items() if k != "threshold_outcomes"}, "thresholds_met": ", ".join(str(x) for x in met)})
            pd.DataFrame(rows).to_excel(writer, sheet_name="By Activity", index=False)

        # Summary sheet: counts per threshold
        summary = aggregated.get("summary", {})
        if summary:
            pd.DataFrame([
                {"threshold_id": k, "count": v} for k, v in summary.items()
            ]).to_excel(writer, sheet_name="Summary", index=False)

        # Optional spec sheet from DOCX
        if spec and spec.get("paragraphs"):
            spec_df = pd.DataFrame({"paragraph": spec["paragraphs"]})
            spec_df.to_excel(writer, sheet_name="Spec", index=False)

    _log.info("Wrote report to %s", path)


def write_persona_report(
    client_reports: list[dict[str, Any]],
    path: str | Path,
    *,
    sheet_name_template: str = "{name}_{age}",
    max_sheet_name_len: int = 31,
    spec: dict[str, Any] | None = None,
) -> None:
    """Write one sheet per client with per-activity 5y/10y statuses."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    summary_rows: list[dict[str, Any]] = []

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for item in client_reports:
            client = item.get("client")
            rows = item.get("rows", [])
            if not client:
                continue

            name = getattr(client, "name", "Client")
            age = getattr(client, "age", None)
            sheet_name = sheet_name_template.format(name=name, age=age or "").strip("_ ").strip()
            sheet_name = _safe_sheet_name(sheet_name, max_sheet_name_len)

            df = pd.DataFrame(rows)
            if df.empty:
                df = pd.DataFrame(
                    columns=[
                        "Activity",
                        "5 Year Critical Failures",
                        "5 Year Supporting Failures",
                        "5 Year Final Status",
                        "10 Year Critical Failures",
                        "10 Year Supporting Failures",
                        "10 Year Final Status",
                    ]
                )
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Summary counts for quick scan
            if not df.empty and "5 Year Final Status" in df.columns and "10 Year Final Status" in df.columns:
                s5 = df["5 Year Final Status"].value_counts(dropna=False).to_dict()
                s10 = df["10 Year Final Status"].value_counts(dropna=False).to_dict()
            else:
                s5 = {}
                s10 = {}
            summary_rows.append(
                {
                    "Client": name,
                    "Age": age,
                    "Gender": getattr(client, "gender", None),
                    "Sheet": sheet_name,
                    "5Y GREEN": int(s5.get("GREEN", 0)),
                    "5Y YELLOW": int(s5.get("YELLOW", 0)),
                    "5Y RED": int(s5.get("RED", 0)),
                    "10Y GREEN": int(s10.get("GREEN", 0)),
                    "10Y YELLOW": int(s10.get("YELLOW", 0)),
                    "10Y RED": int(s10.get("RED", 0)),
                }
            )

        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

        if spec and spec.get("tables"):
            # Flatten rule tables for traceability.
            for i, table in enumerate(spec["tables"][:3]):
                df = pd.DataFrame(table)
                df.to_excel(writer, sheet_name=_safe_sheet_name(f"Logic_{i+1}", max_sheet_name_len), index=False, header=False)

    _log.info("Wrote persona report to %s", path)


def _safe_sheet_name(name: str, max_len: int) -> str:
    invalid = set(r"[]:*?/\\")
    cleaned = "".join("_" if c in invalid else c for c in name)
    cleaned = cleaned.strip()
    if not cleaned:
        cleaned = "Sheet"
    return cleaned[:max_len]
