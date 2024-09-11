"""Módulo de configuração de logs."""

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from rich.console import Console

console = Console()


def setup_logging(log_directory: str = "logs", log_level: int = logging.DEBUG) -> None:
    """Configuração de logs geral.

    Args:
    ----
        log_directory (str, optional): _description_. Defaults to "logs".
        log_level (int, optional): _description_. Defaults to logging.DEBUG.

    """
    if not Path(log_directory).exists():
        Path(log_directory).mkdir(parents=True)

    logger = logging.getLogger()
    logger.setLevel(log_level)  # Global Level's log

    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    # Handler for rotations files
    file_handler = RotatingFileHandler(
        Path(log_directory) / "val_detailed_log.log",
        maxBytes=10**6,
        backupCount=10,
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)

    # Console Handle
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    stream_handler.setLevel(logging.DEBUG)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    logger.debug(console.print("Logging setup complete.", style="bright_yellow"))
