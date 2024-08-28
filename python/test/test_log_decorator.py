"""Teste do log_decorator.py."""

import logging

from src.log_decorator import log_execution


@log_execution
def funcao_teste(a: int, b: int) -> int:
    """Test function.

    Args:
    ----
        a (int): number
        b (int): number

    Returns:
    -------
        int: Sum number

    """
    return a + b


def test_funcao_teste(caplog):
    """Teste da função funcao_teste.

    Args:
    ----
        caplog (_type_): stdout

    """
    with caplog.at_level(logging.DEBUG):
        resultado = funcao_teste(2, 3)
        teste = 5
        assert resultado == teste
        assert "Entrando em funcao_teste com args=(2, 3) e kwargs={}" in caplog.text
        assert "Saindo de funcao_teste com resultado=5" in caplog.text
