import logging
from logging.handlers import RotatingFileHandler
import os


def setup_logging(log_directory="logs", log_level=logging.DEBUG):
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)

    logger = logging.getLogger()
    logger.setLevel(log_level)

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s")

    file_handler = RotatingFileHandler(
        os.path.join(log_directory, "val_detailed_log.log"),
        maxBytes=10**6,
        backupCount=10
    )
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    logger.debug("Logging setup complete.")
