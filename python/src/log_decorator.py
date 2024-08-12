import logging
from functools import wraps
from rich.console import Console


console = Console()
logger = logging.getLogger(__name__)


def log_execution(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        logger.debug(
            console.print(
                f"\nEntrando em {func.__name__} com args={
                    args} e kwargs={kwargs}",
                style="bright_yellow"))
        try:
            result = func(*args, **kwargs)
            logger.info(f"\nSaindo de {func.__name__} com resultado={result}")
            return result
        except Exception as e:
            logger.error(console.print(
                f"\nErro em {func.__name__} com args={args} e kwargs={kwargs}",
                style="bright_yellow"), exc_info=True)
            console.print(
                f"\nErro ao executar a função capturada em log_execution: {e}",
                style="bright_yellow")
    return wrapper
