'''
    Módulo de Testes unitários da aplicação.
'''
# val/src/test/test_main.py

import unittest
import pythoncom
import win32com
import threading
from typing import Type
from rich.console import Console
import src.main
from src.core import val
from src.sapador import down_sap
from src import sap
from src.unitarios.controlador import Controlador
from src.transact_zsbmm216 import Transacao
from src.confere_os import consulta_os

console = Console()


def test_com(gui: Type[win32com.client.CDispatch]) -> None:
    def t_test(sap_id):
        nonlocal gui
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch(
            pythoncom.CoGetInterfaceAndReleaseStream(
                sap_id, pythoncom.IID_IDispatch)
        )
        try:
            # Verifica se o método 'StartTransaction' está presente no objeto CDispatch
            if hasattr(app, "StartTransaction"):
                app.StartTransaction("ZSBMM216")
            else:
                console.print(
                    "[bold red]Método 'StartTransaction' não encontrado no objeto CDispatch[/bold red]")
        except Exception as e:
            console.print(
                f"[bold red]Erro ao chamar StartTransaction: {e}[/bold red]")
        return

    pythoncom.CoInitialize()
    sap_id = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, gui)
    thread = threading.Thread(target=t_test, kwargs={'sap_id': sap_id})
    thread.start()


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
        print(type(gui))
        # test_com(gui)

        # tupla = consulta_os("2320145100", gui, ("4600042888", "340", "100"))
        guia = Transacao("4600043760", "344", "100", gui)
        guia.run_transacao("1234")


if __name__ == '__main__':
    unittest.main()
