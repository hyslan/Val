import logging
from functools import wraps


logger = logging.getLogger(__name__)


def log_execution(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        logger.debug(
            f"Entrando em {func.__name__} com args={args} e kwargs={kwargs}")
        try:
            result = func(*args, **kwargs)
            logger.info(f"Saindo de {func.__name__} com resultado={result}")
            return result
        except Exception as e:
            logger.error(
                f"Erro em {func.__name__} com args={args} e kwargs={kwargs}", exc_info=True)
            print(f"Erro ao executar a função capturada em log_execution: {e}")
    return wrapper
