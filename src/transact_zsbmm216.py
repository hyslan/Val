# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
import threading
from src.sap_connection import connect_to_sap
from src.sap import encerrar_sap

# Adicionando um Lock
lock = threading.Lock()


def run_transacao(ordem, contrato, unadm):
    '''Run thread ZSBMM216
    e faz a transação a transação com o respectivo contrato.'''

    def t_transacao():
        '''Transação preenchida ZSBMM216 - Contrato NOVASP'''
        nonlocal ordem

        # Seção Crítica - uso do Lock
        with lock:
            session = connect_to_sap()
            print("Iniciando valoração.")
            session.StartTransaction("ZSBMM216")
            # Unidade Administrativa
            session.findById("wnd[0]/usr/ctxtP_UND").Text = unadm
            # Contrato
            session.findById("wnd[0]/usr/ctxtP_CONT").Text = contrato
            session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
            sap_ordem = session.findById(
                "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
            sap_ordem.Text = ordem
            session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # Start
    thread = threading.Thread(target=t_transacao)
    thread.start()
    # Aguarde a thread concluir
    thread.join(timeout=300)
    if thread.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()


def novasp(ordem):
    '''Executa thread NOVASP'''
    run_transacao(ordem, "4600041302", "344")


def recape(ordem):
    '''Executa thread RECAPE'''
    run_transacao(ordem, "4600044782", "344")


def gbitaquera(ordem):
    '''Executa thread GB ITAQUERA'''
    run_transacao(ordem, "4600042888", "340")


def nortesul(ordem):
    '''Executa thread NORTE SUL'''
    run_transacao(ordem, "4600043760", "344")
