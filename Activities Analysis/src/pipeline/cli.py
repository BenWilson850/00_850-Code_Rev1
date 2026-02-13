"""CLI for the matrix pipeline."""

import argparse
from pathlib import Path

import yaml

from .core import run_pipeline
from .utils import setup_logging, get_logger


def _load_config(config_path: str | Path | None) -> dict:
    path = Path(config_path or "configs/default.yaml")
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run the persona workbook activity matrix pipeline.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "-c", "--config",
        default="configs/default.yaml",
        help="Path to YAML config file",
    )
    parser.add_argument(
        "-t", "--threshold",
        metavar="PATH",
        help="Path to threshold matrix Excel file (alias for --classifications)",
    )
    parser.add_argument(
        "-a", "--activity",
        metavar="PATH",
        help="Path to activity data Excel file (alias for --limits)",
    )
    parser.add_argument(
        "--clients",
        metavar="PATH",
        help="Path to Client Workbook (persona database) Excel file (required)",
    )
    parser.add_argument(
        "--limits",
        metavar="PATH",
        help="Path to Exercise limits optimised.xlsx",
    )
    parser.add_argument(
        "--classifications",
        metavar="PATH",
        help="Path to Exercise Threshold Classifications.xlsx",
    )
    parser.add_argument(
        "-d", "--docx",
        metavar="PATH",
        help="Path to DOCX specification (optional)",
    )
    parser.add_argument(
        "-o", "--output-dir",
        default="out",
        metavar="DIR",
        help="Output directory for report",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Set log level to DEBUG",
    )
    args = parser.parse_args()

    config = _load_config(args.config)
    if not config.get("paths"):
        config["paths"] = {}
    config["paths"]["client_workbook"] = args.clients or config["paths"].get("client_workbook")
    config["paths"]["limits_matrix"] = args.limits or args.activity or config["paths"].get("limits_matrix")
    config["paths"]["classifications_matrix"] = args.classifications or args.threshold or config["paths"].get("classifications_matrix")
    config["paths"]["docx_spec"] = args.docx or config["paths"].get("docx_spec")
    config["paths"]["output_dir"] = args.output_dir
    if args.verbose and "logging" in config:
        config["logging"]["level"] = "DEBUG"

    setup_logging(config.get("logging", {}))
    log = get_logger(__name__)

    try:
        report_path = run_pipeline(
            config,
            clients_path=args.clients,
            limits_path=args.limits,
            classifications_path=args.classifications,
            threshold_path=args.threshold,
            activity_path=args.activity,
            docx_path=args.docx,
            output_dir=args.output_dir,
        )
        log.info("Done. Report: %s", report_path)
    except Exception as e:
        log.exception("Pipeline failed: %s", e)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
