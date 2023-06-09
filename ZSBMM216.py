import win32com.client , sys 
#Conexão SAP
from SAPConnection import Connect_to_SAP
session = Connect_to_SAP()
# Transação ZSBMM216 Sessão 0
def novasp():
    session.StartTransaction("ZSBMM216")
    session.findById("wnd[0]/usr/ctxtP_UND").Text = "344" # Unidade Administrativa
    session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600041302" # Contrato NOVASP
    session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município