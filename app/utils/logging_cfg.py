"""
Configure logging with Rich console output and optional DEBUG file logging.
"""

from __future__ import annotations

import logging
import logging.handlers
import sys
from typing import Any


def _build_console_handler(*, debug_mode: bool, console: Any = None) -> logging.Handler:
    """
    Build a rich-aware console handler and fall back to stdlib logging if Rich is unavailable.
    """
    try:
        from rich.console import Console
        from rich.logging import RichHandler

        rich_console = console or Console(stderr=True)
        handler = RichHandler(
            console=rich_console,
            show_time=True,
            show_level=True,
            show_path=False,
            rich_tracebacks=debug_mode,
            tracebacks_show_locals=debug_mode,
            markup=True,
            log_time_format="[%d-%b-%Y %H:%M:%S]",
        )
        handler.setLevel(logging.INFO)
        return handler
    except Exception:
        handler = logging.StreamHandler(stream=sys.stderr)
        handler.setLevel(logging.INFO)
        fallback_formatter = logging.Formatter(
            fmt="[%(asctime)s] %(levelname)s: %(message)s",
            datefmt="%d-%b-%Y %H:%M:%S",
        )
        handler.setFormatter(fallback_formatter)
        return handler


def configure_logging(debug_mode: bool, *, console: Any = None) -> None:
    """
    If debug_mode is True:
      - root logger level = DEBUG
      - create file handler at DEBUG
      - console at INFO
    Otherwise:
      - root logger level = INFO
      - no file is created
      - console at INFO
    """
    logger = logging.getLogger()
    # Remove any existing handlers if re-run
    logger.handlers.clear()

    if debug_mode:
        logger.setLevel(logging.DEBUG)
        # File handler at DEBUG
        fh = logging.handlers.RotatingFileHandler(
            "batch_bulk_editor.log",
            mode="a",
            maxBytes=10_000_000,
            backupCount=10,
            encoding="utf-8",
        )
        fh.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter(
            fmt="[%(asctime)s] %(levelname)s\t[%(name)s.%(funcName)s:%(lineno)d]\t%(message)s",
            datefmt="%d-%b-%Y %H:%M:%S",
        )
        fh.setFormatter(file_formatter)
        logger.addHandler(fh)

        logger.addHandler(_build_console_handler(debug_mode=True, console=console))
    else:
        # No file handler, just console
        logger.setLevel(logging.INFO)
        logger.addHandler(_build_console_handler(debug_mode=False, console=console))
