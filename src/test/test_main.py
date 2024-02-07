'''
    Módulo de Testes unitários da aplicação.
'''
# val/src/test/test_main.py

import unittest
from unittest.mock import Mock, patch
from rich.console import Console
import src.main
from src.core import val
from src.sapador import down_sap
from src import sap
from src.transact_zsbmm216 import Transacao
from src.unitarios.ligacao_agua.m_ligacao_agua import LigacaoAgua
console = Console()


class TestMain(unittest.TestCase):
    '''Class representing for testing'''

    def test_core(self):
        '''Test main core module loop'''
        result = val(["2341816055"], "4600041302", "344")
        correto = "2341816055", 1, True
        self.assertEqual(result, correto, "Resultado não esperado")

    def test_main_contratada(self):
        '''Test fuction of main module loop'''
        self.assertLogs(src.main.contratada(), level='DEBUG')

    def test_sapador(self):
        '''Test download sapgui'''
        self.assertLogs(down_sap(), level='DEBUG')

    def test_sap(self):
        '''Test COM SapGui module'''
        sapgui = sap.Sap()
        self.assertLogs(sapgui.listar_conexoes(), level='DEBUG')
        sessions = sapgui.listar_sessoes()
        self.assertGreater(len(sessions), 0, "Nenhuma sessão encontrada.")
        self.assertLogs(sapgui.contar_sessoes(), level='DEBUG')
        self.assertLogs(sapgui.criar_sessao(sessions), level='DEBUG')
        n = self.assertLogs(sapgui.escolher_sessao(), level='DEBUG')
        console.print([f"[italic]{n}"])

    def test_dicionario_un(self):
        '''Teste do dicionário com classe'''
        sessao = sap.Sap()
        gui = sessao.escolher_sessao()
        # seletor = Controlador("134000", None, None, None, 1,
        #                       None, "cavalete", None, None, gui)
        # seletor.executar_processo()

    def test_transacao(self):
        '''teste da ZSBMM216'''
        sessao = sap.Sap()
        gui = sessao.escolher_sessao()
        guia = Transacao("4600043760", "344", "100", gui)
        guia.run_transacao("1234")


class TestMetodoPosicaoPagar(unittest.TestCase):
    '''Teste do método _posicao_pagar da classe LigacaoAgua'''

    def setUp(self):
        # substitua pelo construtor real da classe
        self.objeto = LigacaoAgua()
        self.objeto._ramal = False
        self.objeto.session = Mock()
        self.objeto.preco = Mock()
        self.objeto.preco.modifyCell = Mock()
        self.objeto.preco.setCurrentCell = Mock()
        self.objeto.preco.pressEnter = Mock()

    @patch('builtins.print')
    def test_posicao_pagar(self, mock_print):
        '''Testa o método _posicao_pagar'''
        preco_tse = 'preco_tse_teste'
        self.objeto._posicao_pagar(preco_tse)
        self.objeto.preco.modifyCell.assert_called_once_with(
            self.objeto.preco.CurrentCellRow, "QUANT", "1")
        self.objeto.preco.setCurrentCell.assert_called_once_with(
            self.objeto.preco.CurrentCellRow, "QUANT")
        self.objeto.preco.pressEnter.assert_called_once()
        mock_print.assert_called_once_with(f"Pago 1 UN de {preco_tse}")
        self.assertTrue(self.objeto._ramal)


if __name__ == '__main__':
    unittest.main()


if __name__ == '__main__':
    unittest.main()
