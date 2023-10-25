# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
import threading
from sap_connection import connect_to_sap
from sap import encerrar_sap

# Adicionando um Lock
lock = threading.Lock()


def novasp(ordem):
    '''Run thread NOVAS SP'''

    def t_novasp():
        '''Transação preenchida ZSBMM216 - Contrato NOVASP'''
        nonlocal ordem

        # Seção Crítica - uso do Lock
        with lock:
            session = connect_to_sap()
            print("Iniciando valoração.")
            session.StartTransaction("ZSBMM216")
            # Unidade Administrativa
            session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
            # Contrato NOVASP
            session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600041302"
            session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
            sap_ordem = session.findById(
                "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
            sap_ordem.Text = ordem
            session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # Start thread novasp.
    thread = threading.Thread(target=t_novasp)
    thread.start()
    # Timeout 5min
    thread.join(timeout=300)
    if thread.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()


def recape(ordem):
    '''Run thread RECAPE'''

    async def t_recape():
        '''Transação preenchida ZSBMM216 - Contrato RECAPE'''
        nonlocal ordem

        # Seção Crítica - uso do Lock
        with lock:
            session = connect_to_sap()
            print("Iniciando valoração.")
            session.StartTransaction("ZSBMM216")
            # Unidade Administrativa
            session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
            # Contrato RECAPE
            session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600044782"
            session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
            sap_ordem = session.findById(
                "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
            sap_ordem.Text = ordem
            session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # Start thread recape.
    thread = threading.Thread(target=t_recape)
    thread.start()
    # Timeout 5min
    thread.join(timeout=300)
    if thread.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()
