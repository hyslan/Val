"""Decorator para logar a execução de funções."""

import logging
from functools import wraps
from typing import Any

from rich.console import Console

console = Console()
logger = logging.getLogger(__name__)


def log_execution(func) -> Any:
    """Log the execution of functions.

    Args:
    ----
        func: The function to be decorated.

    Returns:
    -------
        The decorated function.

    """

    @wraps(func)
    def wrapper(*args, **kwargs) -> Any:
        logger.debug(
            console.print(
                f"\nEntrando em {func.__name__} com args={
                    args} e kwargs={kwargs}",
                style="bright_yellow",
            ),
        )
        try:
            result = func(*args, **kwargs)
            logger.info("\nSaindo de %s com resultado=%s", func.__name__, result)
        except Exception as e:
            logger.exception(
                console.print(f"\nErro em {func.__name__} com args={args} e kwargs={kwargs}", style="bright_yellow"),
            )
            console.print(f"\nErro ao executar a função capturada em log_execution: {e}", style="bright_yellow")
        else:
            return result

    return wrapper
