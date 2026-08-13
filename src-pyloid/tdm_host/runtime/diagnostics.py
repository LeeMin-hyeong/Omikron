"""Local diagnostic logging that never exposes tracebacks through RPC."""

from __future__ import annotations

import logging
import os
import tempfile
import uuid
from logging.handlers import RotatingFileHandler
from pathlib import Path


_LOGGER_NAME = "tdm"


def diagnostic_log_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
    return base / "TDM" / "logs" / "tdm.log"


def _diagnostic_logger() -> logging.Logger:
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    logger.propagate = False
    formatter = logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    try:
        path = diagnostic_log_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        handler: logging.Handler = RotatingFileHandler(
            path,
            maxBytes=2 * 1024 * 1024,
            backupCount=5,
            encoding="utf-8",
        )
    except OSError:
        handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger


def record_exception(exc: BaseException, *, context: str = "RPC") -> str:
    """Write the complete exception locally and return a short correlation ID."""
    error_id = uuid.uuid4().hex[:12].upper()
    _diagnostic_logger().error(
        "[%s] %s: %s",
        error_id,
        context,
        exc,
        exc_info=(type(exc), exc, exc.__traceback__),
    )
    return error_id


def diagnostic_detail(exc: BaseException, *, context: str = "RPC") -> str:
    return f"오류 ID: {record_exception(exc, context=context)}"
