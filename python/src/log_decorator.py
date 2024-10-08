"""Decorator para logar a execução de funções."""

import logging
from collections.abc import Callable
from functools import wraps
from typing import Any

logger = logging.getLogger(__name__)


def log_execution(func: Callable[..., Any]) -> Callable[..., Any]:
    """Log the execution of functions.

    Args:
    ----
        func: The function to be decorated.

    Returns:
    -------
        The decorated function.

    """

    @wraps(func)
    def wrapper(*args: tuple[Any, ...]) -> Any:
        logger.debug("Entrando em %s com args=%s", func.__name__, args)
        try:
            result = func(*args)
            logger.info("Saindo de %s com resultado=%s", func.__name__, result)
        except Exception:
            logger.exception("Erro em %s com args=%s", func.__name__, args)
        else:
            return result

    return wrapper
