"""Decorator para logar a execução de funções."""

import logging
from functools import wraps
from typing import Any

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
        logger.debug("Entrando em %s com args=%s e kwargs=%s", func.__name__, args, kwargs)
        try:
            result = func(*args, **kwargs)
            logger.info("Saindo de %s com resultado=%s", func.__name__, result)
        except Exception:
            logger.exception("Erro em %s com args=%s e kwargs=%s", func.__name__, args, kwargs)
        else:
            return result

    return wrapper
