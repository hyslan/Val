# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
from sap_connection import connect_to_sap
session = connect_to_sap()


def novasp(ordem):
    '''Transação preenchida ZSBMM216 - Sessão 0'''
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
