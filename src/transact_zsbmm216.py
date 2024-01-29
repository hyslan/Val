# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
import threading
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

        def t_transacao():
            '''Transação preenchida ZSBMM216 - Contrato NOVASP'''
            nonlocal ordem

            # Seção Crítica - uso do Lock
            with lock:
                print("Iniciando valoração.")
                self.session.StartTransaction("ZSBMM216")
                # Unidade Administrativa
                self.session.findById("wnd[0]/usr/ctxtP_UND").Text = self.unadm
                # Contrato
                self.session.findById(
                    "wnd[0]/usr/ctxtP_CONT").Text = self.contrato
                self.session.findById(
                    "wnd[0]/usr/ctxtP_MUNI").Text = self.municipio  # Cidade
                sap_ordem = self.session.findById(
                    "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
                sap_ordem.Text = ordem
                self.session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

        # Start
        thread = threading.Thread(target=t_transacao)
        thread.start()
        # Aguarde a thread concluir
        thread.join(timeout=300)
        if thread.is_alive():
            print("SAP demorando mais que o esperado, encerrando.")
            sap.encerrar_sap()
