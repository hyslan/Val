"""Módulo de testes do módulo core."""

from unittest.mock import MagicMock, patch

import pandas as pd
import pytest

from src.core import val


@pytest.fixture
def mock_session() -> MagicMock:
    """Mock session for testing purposes."""
    return MagicMock()


# Dados de teste
df_test = pd.DataFrame([["2425188908", "100"]], index=[0], columns=["Ordem", "COD_MUNICIPIO"])
pendentes_array = df_test.to_numpy()
# Parâmetros de teste
revalorar = False
contrato = "4600041302"
token = ""
n_con = 0


@patch("src.confere_os.pythoncom.CoMarshalInterThreadInterfaceInStream")
@patch("src.confere_os.pythoncom.CoGetInterfaceAndReleaseStream")
@patch("src.confere_os.win32.Dispatch")
def test_core_val(mock_dispatch, mock_get_interface, mock_marshal, capsys, mock_session: MagicMock) -> None:
    """Teste do módulo core.val.

    Args:
    ----
        mock_session (MagicMock): Mock session for testing purposes.

    """
    # Configurar o mock para CoMarshalInterThreadInterfaceInStream
    mock_marshal.return_value = MagicMock()
    # Configurar o mock para CoGetInterfaceAndReleaseStream
    mock_get_interface.return_value = MagicMock()
    # Configurar o mock para Dispatch
    mock_dispatch.return_value = MagicMock()

    # Test case 1: Valid input
    val(pendentes_array, session=mock_session, contrato=contrato, revalorar=revalorar, token=token, n_con=n_con)
    # Capturar a Saída
    captured = capsys.readouterr()
    print(captured.out)

    # Test case 2: Empty input
    msg = "Nenhuma ordem para ser valorada."
    with pytest.raises(ValueError, match=msg):
        val(pendentes_array=[], session=mock_session, contrato=None, revalorar=False, token="", n_con=0)
