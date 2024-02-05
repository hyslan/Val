# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
import threading
import pythoncom
import win32com.client as win32
from src.sap import Sap

# Adicionando um Lock
lock = threading.Lock()
sap = Sap()


class Transacao():
    '''Classe operadora da transação 216'''

    def __init__(self, contrato, unadm,
                 municipio, session) -> None:
        self.contrato = contrato
        self.unadm = unadm
        self.session = session
        self.municipio = municipio

    def run_transacao(self, ordem):
        '''Run thread ZSBMM216
        e faz a transação a transação com o respectivo contrato.'''

        def t_transacao(session_id):
            '''Transação preenchida ZSBMM216 - Contrato NOVASP'''
            nonlocal ordem

            # Seção Crítica - uso do Lock
            with lock:
                # pylint: disable=E1101
                pythoncom.CoInitialize()
                # pylint: disable=E1101
                gui = win32.Dispatch(
                    pythoncom.CoGetInterfaceAndReleaseStream(
                        session_id, pythoncom.IID_IDispatch)
                )
                print("Iniciando valoração.")
                gui.StartTransaction("ZSBMM216")
                # Unidade Administrativa
                gui.findById("wnd[0]/usr/ctxtP_UND").Text = self.unadm
                # Contrato
                gui.findById(
                    "wnd[0]/usr/ctxtP_CONT").Text = self.contrato
                gui.findById(
                    "wnd[0]/usr/ctxtP_MUNI").Text = self.municipio  # Cidade
                sap_ordem = gui.findById(
                    "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
                sap_ordem.Text = ordem
                gui.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

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
            # sap.encerrar_sap()
