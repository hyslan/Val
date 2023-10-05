# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
import threading
from sap_connection import connect_to_sap
from sap import fechar_conexao


def novasp(ordem):
    '''Threading NOVAS SP'''

    def t_novasp():
        '''Transação preenchida ZSBMM216 - Contrato NOVASP'''
        nonlocal ordem
        session = connect_to_sap()
        print("Iniciando valoração.")
        session.StartTransaction("ZSBMM216")
        # Unidade Administrativa
        session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
        # Contrato NOVASP
        session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600041302"
        session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
        sap_ordem = session.findById("wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
        sap_ordem.Text = ordem
        session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # Thread para função aninhada.
    thread_novasp = threading.Thread(target=t_novasp)
    thread_novasp.start()
    # Aguardar até o limite de 5min.
    thread_novasp.join(timeout=300)
    # Verificar se está em execução.
    if thread_novasp.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        fechar_conexao()


def recape(ordem):
    '''Threading RECAPE'''

    def t_recape():
        '''Transação preenchida ZSBMM216 - Contrato RECAPE'''
        nonlocal ordem
        session = connect_to_sap()
        print("Iniciando valoração.")
        session.StartTransaction("ZSBMM216")
        # Unidade Administrativa
        session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
        # Contrato RECAPE
        session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600044782"
        session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
        sap_ordem = session.findById("wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
        sap_ordem.Text = ordem
        session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # Thread para função aninhada.
    thread_recape = threading.Thread(target=t_recape)
    thread_recape.start()
    # Aguardar até o limite de 5min.
    thread_recape.join(timeout=300)
    # Verificar se está em execução.
    if thread_recape.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        fechar_conexao()
