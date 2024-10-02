"""Módulo de testes para o módulo sap.py."""

from python.src import sap
from python.src.sapador import down_sap


def setup() -> str:
    """Set up the test environment.

    Returns:
        str: TOKEN SSO

    """
    return down_sap()


def test_session_active() -> None:
    """Test if the session is active."""
    token = setup()
    n_con = sap.count_connections()
    sap.keep_session_active(n_con, token)
