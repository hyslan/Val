# transact_zsbmm216.py
"""Módulo de Contrato"""
# Conexão SAP
import threading
import pythoncom
import pywintypes
import win32com.client as win32
from rich.console import Console
from src import sap

# Adicionando um Lock
lock = threading.Lock()
console = Console()


class Transacao:
    """Classe operadora da transação 216"""

    def __init__(self, contrato,
                 municipio, session) -> None:
        self._contrato = contrato
        self.session = session
        self._municipio = municipio

    @property
    def municipio(self):
        """Getter para municipio"""
        return self._municipio

    @municipio.setter
    def municipio(self, cod):
        """Setter for municipio"""
        if isinstance(cod, str):
            self._municipio = cod
        else:
            raise ValueError("Wrong type, need to be string.")

    @property
    def contrato(self):
        return self._contrato

    @contrato.setter
    def contrato(self, cod):
        if isinstance(cod, str):
            self._contrato = cod
        else:
            raise ValueError("Wrong type, need to be string.")

    def run_transacao(self, ordem, tipo="individual"):
        """Run thread ZSBMM216
        e faz a transação a transação com o respectivo contrato."""

        def t_transacao(session_id):
            """Transação preenchida ZSBMM216 - Contrato NOVASP"""
            nonlocal ordem

            # Seção Crítica - uso do Lock
            with lock:
                try:
                    # pylint: disable=E1101
                    pythoncom.CoInitialize()
                    # pylint: disable=E1101
                    gui = win32.Dispatch(
                        pythoncom.CoGetInterfaceAndReleaseStream(
                            session_id, pythoncom.IID_IDispatch)
                    )
                    print("Iniciando Transação ZSBM216.")
                    gui.StartTransaction("ZSBMM216")
                    if tipo == "consulta":
                        gui.findById("wnd[0]/usr/radRB_CON").Select()

                    # Unidade Administrativa
                    gui.findById("wnd[0]/usr/ctxtP_UND").Text = "1093"
                    # Contrato
                    gui.findById(
                        "wnd[0]/usr/ctxtP_CONT").Text = self.contrato
                    gui.findById(
                        "wnd[0]/usr/ctxtP_MUNI").Text = self._municipio  # Cidade
                    sap_ordem = gui.findById(
                        "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
                    sap_ordem.Text = ordem
                    gui.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

                except pywintypes.com_error as erro_216:
                    console.print(f"Erro na ZSBMM216: {erro_216}")
                    console.print_exception(show_locals=True)

        try:
            # pylint: disable=E1101
            pythoncom.CoInitialize()
            session_id = pythoncom.CoMarshalInterThreadInterfaceInStream(
                pythoncom.IID_IDispatch, self.session)
            # Start
            thread = threading.Thread(target=t_transacao, kwargs={
                'session_id': session_id})
            thread.start()
            # Aguarde a thread concluir
            thread.join(timeout=300)
            if thread.is_alive():
                print("SAP demorando mais que o esperado, encerrando.")
                sap.encerrar_sap()

        except pywintypes.com_error as erro_thread216:
            console.print(f"Erro na thread da ZSBMM216: {erro_thread216}")
            console.print_exception(show_locals=True)
