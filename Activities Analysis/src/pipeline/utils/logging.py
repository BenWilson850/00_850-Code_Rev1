"""Logging setup for the pipeline."""

import logging
import sys
from pathlib import Path
from typing import Any

_configured = False
_loggers: dict[str, logging.Logger] = {}


def setup_logging(config: dict[str, Any] | None = None) -> None:
    """Configure root/pipeline logging from config.

    Config keys: level (DEBUG, INFO, WARNING, ERROR), format, file (optional path).
    """
    global _configured
    if _configured:
        return
    config = config or {}
    level_name = (config.get("level") or "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)
    fmt = config.get("format") or "%(asctime)s [%(levelname)s] %(name)s: %(message)s"

    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    log_file = config.get("file")
    if log_file:
        Path(log_file).parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    logging.basicConfig(
        level=level,
        format=fmt,
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
        force=True,
    )
    # Reduce noise from third-party libs
    logging.getLogger("openpyxl").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    _configured = True


def get_logger(name: str) -> logging.Logger:
    """Return a logger for the given module name (e.g. __name__)."""
    if name not in _loggers:
        _loggers[name] = logging.getLogger(name)
    return _loggers[name]
