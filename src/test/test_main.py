'''
    Módulo de Testes unitários da aplicação.
'''
# val/src/test/test_main.py

import unittest
from rich.console import Console
import src.main
from src.core import val
from src.sapador import down_sap
from src import sap

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


if __name__ == '__main__':
    unittest.main()
