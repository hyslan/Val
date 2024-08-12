import logging
from functools import wraps


logger = logging.getLogger(__name__)


def log_execution(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        logger.debug(
            f"Entrando em {func.__name__} com args={args} e kwargs={kwargs}")
        result = func(*args, **kwargs)
        logger.debug(f"Saindo de {func.__name__} com resultado={result}")
        return result
    return wrapper
