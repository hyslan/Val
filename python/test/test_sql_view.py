"""Módulo de testes de SQL View.

Conexões e queries.
"""

from unittest.mock import MagicMock, patch

import numpy as np
import pytest

from python.src.sql_view import Sql


@pytest.fixture
def sql_instance() -> Sql:
    """Fixture para Sql.

    Returns
    -------
        Sql: Instância de Sql.

    """
    with patch("python.src.sql_view.sa.create_engine") as mock_engine:
        mock_engine.return_value.connect.return_value = MagicMock()
        return Sql(ordem="12345", cod_tse="67890")


def test_ordem_property(sql_instance: Sql) -> None:
    """Teste da propriedade ordem.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    assert sql_instance.ordem == "12345"
    sql_instance.ordem = "54321"
    assert sql_instance.ordem == "54321"


def test_check_employee(sql_instance: Sql) -> None:
    """Teste do método __check_employee.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_string.return_value = "John Doe"
        result = sql_instance._Sql__check_employee("117615")  # type: ignore # noqa: SLF001
        assert result == "John Doe"


def test_carteira_tse(sql_instance: Sql) -> None:
    """Teste do método carteira_tse.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.carteira_tse("contract123", ["tse1", "tse2"])
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_tse_escolhida(sql_instance: Sql) -> None:
    """Teste do método tse_escolhida.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql, patch("builtins.input", side_effect=["n"]):
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.tse_escolhida("contract123")
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_tse_expecifica(sql_instance: Sql) -> None:
    """Teste do método tse_expecifica.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql, patch("builtins.input", side_effect=["n"]):
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.tse_expecifica("contract123")
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_clean_duplicates(sql_instance: Sql) -> None:
    """Teste do método Clean Duplicates.

    Args:
    ----
        sql_instance (Sql): Instância de Sql._

    """
    with patch("python.src.sql_view.sa.create_engine") as mock_engine:
        mock_cnn = mock_engine.return_value.connect.return_value
        sql_instance.clean_duplicates()
        mock_cnn.execute.assert_called_once()
        mock_cnn.commit.assert_called_once()
        mock_cnn.close.assert_called_once()


def test_retrabalho_search(sql_instance: Sql) -> None:
    """Teste do método retrabalho_search.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.retrabalho_search("2023-01-01", "2023-12-31")
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_valorada(sql_instance: Sql) -> None:
    """Teste do método valorada.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with (
        patch("python.src.sql_view.sa.create_engine") as mock_engine,
        patch("python.src.sql_view.Sql._Sql__check_employee", return_value="John Doe"),
    ):
        mock_cnn = mock_engine.return_value.connect.return_value
        sql_instance.valorada(
            valorado="Sim",
            contrato="contract123",
            municipio="municipio123",
            status="status123",
            obs="obs123",
            data_valoracao="01-01-2023",
            matricula="123456",
            valor_medido=100.0,
            tempo_gasto=2.0,
        )
        mock_cnn.execute.assert_called_once()
        mock_cnn.commit.assert_called_once()
        mock_cnn.close.assert_called_once()


def test_ordem_especifica(sql_instance: Sql) -> None:
    """Teste de Ordem Específica.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.ordem_especifica("contract123")
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_familia(sql_instance: Sql) -> None:
    """Teste de Família.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.familia(["family1", "family2"], "contract123")
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_desobstrucao(sql_instance: Sql) -> None:
    """Teste de Desobstrução.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_numpy.return_value = np.array([[1, 2], [3, 4]])
        result = sql_instance.desobstrucao()
        assert (result == np.array([[1, 2], [3, 4]])).all()


def test_show_family(sql_instance: Sql) -> None:
    """Teste de Show Family.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value["FAMILIA"].to_string.return_value = "family1\nfamily2"
        result = sql_instance.show_family()
        assert result == "family1\nfamily2"


def test_show_tses(sql_instance: Sql) -> None:
    """Teste de Show TSEs.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_string.return_value = "tse1\ntse2"
        result = sql_instance.show_tses()
        assert result == "tse1\ntse2"


def test_get_new_hidro(sql_instance: Sql) -> None:
    """Teste de Get New Hidro.

    Args:
    ----
        sql_instance (Sql): Instância de Sql.

    """
    with patch("python.src.sql_view.pd.read_sql") as mock_read_sql:
        mock_read_sql.return_value.to_string.return_value = "hidro123"
        result = sql_instance.get_new_hidro()
        assert result == "hidro123"
