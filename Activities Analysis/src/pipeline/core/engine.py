"""Pipeline engine: orchestrates read → normalize → assess → aggregate → write."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..utils import setup_logging, get_logger
from .persona_pipeline import run_persona_pipeline


def run_pipeline(
    config: dict[str, Any],
    *,
    clients_path: str | Path | None = None,
    limits_path: str | Path | None = None,
    classifications_path: str | Path | None = None,
    # Back-compat aliases (older CLI flags)
    threshold_path: str | Path | None = None,  # classifications matrix
    activity_path: str | Path | None = None,  # limits matrix
    docx_path: str | Path | None = None,
    output_dir: str | Path | None = None,
) -> Path:
    """Run the persona workbook activity matrix pipeline."""
    setup_logging(config.get("logging", {}))
    log = get_logger(__name__)

    paths = config.get("paths", {})
    clients_path = clients_path or paths.get("client_workbook")
    limits_path = limits_path or activity_path or paths.get("limits_matrix")
    classifications_path = classifications_path or threshold_path or paths.get("classifications_matrix")
    docx_path = docx_path or paths.get("docx_spec")
    output_dir = Path(output_dir or paths.get("output_dir", "out"))

    if not clients_path:
        raise ValueError("client_workbook path is required (--clients)")
    if not limits_path or not classifications_path:
        raise ValueError("limits_matrix and classifications_matrix paths are required")

    output_dir.mkdir(parents=True, exist_ok=True)
    log.info("Running persona activity matrix report")
    return run_persona_pipeline(
        config,
        clients_path=clients_path,
        limits_path=limits_path,
        classifications_path=classifications_path,
        docx_path=docx_path,
        output_dir=output_dir,
    )
