'''Módulo de Contrato'''
# Conexão SAP
from sap_connection import connect_to_sap
session = connect_to_sap()
# Transação ZSBMM216 Sessão 0


def novasp():
    '''Transação preenchida ZSBMM216'''
    session.StartTransaction("ZSBMM216")
    # Unidade Administrativa
    session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
    # Contrato NOVASP
    session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600041302"
    session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
