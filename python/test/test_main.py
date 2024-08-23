"""Módulo de Testes unitários da aplicação."""
import argparse

# val/src/test/test_main.py
import unittest
from unittest.mock import Mock, patch

import numpy as np
from rich.console import Console

import main
from python.src import sap
from python.src.confere_os import consulta_os
from python.src.core import val
from python.src.sapador import down_sap
from python.src.transact_zsbmm216 import Transacao
from python.src.unitarios.controlador import Controlador
from python.src.unitarios.ligacao_agua.m_ligacao_agua import LigacaoAgua
from python.src.unitarios.poco.m_poco import Poco
from python.src.wms.consulta_estoque import estoque

console = Console()


class TestMain(unittest.TestCase):
    """Class representing for testing."""

    @patch("argparse.ArgumentParser.parse_args")
    def test_main(self, mock_parse_args) -> None:
        """Test main function."""
        # mock_args = Mock(return_value=("0", "9", "4600041302", "alefafa"))
        mock_parse_args.return_value = argparse.Namespace(
            s="0", o="9", c="4600041302", p="alefafa")
        assert main.main() is None

    def test_core(self) -> None:
        """Test main core module loop."""
        arr = np.array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10])
        val(arr, "0", "4600041302", False)

    def test_estoque(self) -> None:
        """Teste de estoque."""
        estoque(
            sap.choose_connection(0), sap.listar_sessoes(), ("4600041302", "344", "100"))

    def test_consulta(self) -> None:
        """Consulta de OS no SAP."""
        consulta_os(
            "2400341804", sap.choose_connection(0), ("4600041302", "344", "100"))

    def test_sapador(self) -> None:
        """Test download sapgui."""
        self.assertLogs(down_sap(), level="DEBUG")

    def test_sap(self) -> None:
        """Test COM SapGui module."""
        self.assertLogs(sap.connection_object(0), level="DEBUG")
        sessions = sap.listar_sessoes()
        assert len(sessions) > 0, "Nenhuma sessão encontrada."
        self.assertLogs(sap.contar_sessoes(), level="DEBUG")
        self.assertLogs(sap.create_session(sessions), level="DEBUG")
        n = self.assertLogs(sap.choose_connection(), level="DEBUG")
        console.print([f"[italic]{n}"])

    def test_dicionario_un(self) -> None:
        """Teste do dicionário com classe."""
        gui = sap.choose_connection(0)
        seletor = Controlador("134000", None, None, None, 1,
                              None, "cavalete", None, None, gui)
        seletor.executar_processo()

    def test_transacao(self) -> None:
        """Teste da ZSBMM216."""
        gui = sap.choose_connection()
        guia = Transacao("4600043760", "344", "100", gui)
        guia.run_transacao("1234")


class TestMetodoPosicaoPagar(unittest.TestCase):
    """Teste do método _posicao_pagar da classe LigacaoAgua."""

    def setUp(self) -> None:
        self.objeto = LigacaoAgua(*Mock())
        self.objeto._ramal = False
        self.objeto.session = Mock()
        self.objeto.preco = Mock()
        self.objeto.preco.modifyCell = Mock()
        self.objeto.preco.setCurrentCell = Mock()
        self.objeto.preco.pressEnter = Mock()

    @patch("builtins.print")
    def test_posicao_pagar(self, mock_print) -> None:
        """Testa o método _posicao_pagar."""
        preco_tse = "preco_tse_teste"
        self.objeto._posicao_pagar(preco_tse)
        self.objeto.preco.modifyCell.assert_called_once_with(
            self.objeto.preco.CurrentCellRow, "QUANT", "1")
        self.objeto.preco.setCurrentCell.assert_called_once_with(
            self.objeto.preco.CurrentCellRow, "QUANT")
        self.objeto.preco.pressEnter.assert_called_once()
        mock_print.assert_called_once_with(f"Pago 1 UN de {preco_tse}")
        assert self.objeto._ramal


class TestProcessarOperacao(unittest.TestCase):
    """Testar classe poço."""

    def setUp(self) -> None:
        gui = sap.choose_connection()
        self.objeto = Poco(
            etapa="2400145264",
            corte=None,
            relig=None,
            reposicao=None,
            num_tse_linhas=1,
            etapa_reposicao=None,
            identificador="poço",
            posicao_rede=None,
            profundidade=None,
            session=gui,
        )
        # self.objeto._pagar = Mock()

    @patch("builtins.print")
    def test_processar_operacao_nivelamento(self, mock_print) -> None:
        """Niv pv."""
        tipo_operacao = "NIVELAMENTO"
        self.objeto.CODIGOS.get(tipo_operacao)
        self.objeto._processar_operacao(tipo_operacao)
        mock_print.assert_called_once_with(
            f"Iniciando processo de pagar {tipo_operacao.replace('_', ' ')}")

    # self.objeto._pagar.assert_called_with(codigo[1])

    def case_test(self) -> None:
        """Teste de caso."""
        self.objeto.nivelamento()


if __name__ == "__main__":
    unittest.main()
