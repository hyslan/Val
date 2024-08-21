import logging
import os
from logging.handlers import RotatingFileHandler

from rich.console import Console

console = Console()


def setup_logging(log_directory="logs", log_level=logging.DEBUG):
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)

    logger = logging.getLogger()
    logger.setLevel(log_level)  # Global Level's log

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    # Handler for rotations files
    file_handler = RotatingFileHandler(
        os.path.join(log_directory, "val_detailed_log.log"),
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

    logger.debug(console.print(
        "Logging setup complete.", style="bright_yellow"))
