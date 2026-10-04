"""
Logging configuration module for the AWS BAU Report application.

Provides structured logging with both console and file output.
"""

import logging
import os
import sys
from typing import Optional


def setup_logger(
    name: str = "bau_report",
    log_level: str = "INFO",
    log_file: Optional[str] = None
) -> logging.Logger:
    """
    Configure and return a structured logger with console and optional file handlers.

    Args:
        name: Logger name (default: 'bau_report')
        log_level: Logging level string (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Optional path to log file

    Returns:
        Configured logging.Logger instance
    """
    logger = logging.getLogger(name)

    # Prevent duplicate handlers if logger already configured
    if logger.handlers:
        return logger

    level = getattr(logging, log_level.upper(), logging.INFO)
    logger.setLevel(level)

    # Log format
    formatter = logging.Formatter(
        fmt="[%(asctime)s] %(levelname)-8s %(name)-20s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console handler (stdout)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler (if specified)
    if log_file:
        try:
            log_dir = os.path.dirname(log_file)
            if log_dir and not os.path.exists(log_dir):
                os.makedirs(log_dir, exist_ok=True)

            file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
            file_handler.setLevel(level)
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except (OSError, PermissionError) as e:
            logger.warning(f"Could not create log file '{log_file}': {e}")

    return logger


def get_logger(name: str = "bau_report") -> logging.Logger:
    """
    Get an existing logger by name. If not configured, returns a basic logger.

    Args:
        name: Logger name

    Returns:
        logging.Logger instance
    """
    return logging.getLogger(name)
