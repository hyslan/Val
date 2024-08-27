"""Módulo de testes do módulo core."""

from unittest.mock import MagicMock

import pandas as pd
import pytest

from src.core import val


@pytest.fixture
def mock_session(mocker) -> MagicMock:
    """Mock session for testing purposes."""
    mock_get_object = mocker.patch("src.core.win32com.client.GetObject")
    mock_sap_gui = MagicMock()
    mock_get_object.return_value.GetScriptingEngine.return_value.Connections = "mock_connections"
    return mock_sap_gui


df_test = pd.DataFrame([["2425188908", "100"]], index=[0], columns=["Ordem", "COD_MUNICIPIO"])
pendentes_array = df_test.to_numpy()


def test_core_val(capsys, mock_session: MagicMock) -> None:
    """Teste do módulo core.val.

    Args:
    ----
        mock_session (MagicMock): Mock session for testing purposes.

    """
    revalorar = False
    contrato = "4600041302"
    token = ""
    n_con = 0
    # Test case 1: Valid input
    val(pendentes_array, session=mock_session, contrato=contrato, revalorar=revalorar, token=token, n_con=n_con)
    # Capturar a Saída
    captured = capsys.readouterr()
    print(captured.out)

    # Test case 2: Empty input
    msg = "Nenhuma ordem para ser valorada."
    with pytest.raises(ValueError, match=msg):
        val(pendentes_array=0, session=mock_session, contrato=None, revalorar=False, token="", n_con=0)
