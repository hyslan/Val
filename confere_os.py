# ConfereOS.py
'''Módulo de check-up no SAP'''
from sap_connection import connect_to_sap

def consulta_os(n_os):
    '''Função para consultar ORDEM na transação ZSBPM020.'''
    session = connect_to_sap()
    session.StartTransaction("ZSBPM020")
    status_usuario = "USTXT"
    status_sistema = "STTXT"
    campo_os = session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW")
    campo_os.Text = n_os
    session.findById("wnd[0]/usr/txtS_CONTR-LOW").text = "4600041302"
    session.findById("wnd[0]/usr/txtS_UN_ADM-LOW").text = "344"
    session.findById("wnd[0]").sendVKey(8)
    consulta = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    status_sistema = consulta.GetCellValue(0, "STTXT")
    status_usuario = consulta.GetCellValue(0, "USTXT")
    corte = consulta.GetCellValue(0, "ZZLOCAL_Corte")  # Supressão
    relig = consulta.GetCellValue(0, "ZZLOCAL_ReligA")  # religação
    posicao_rede = consulta.GetCellValue(
        0, "ZZPOSICAO_REDE")  # Posição da Rede
    profundidade = consulta.GetCellValue(0, "ZZProfundidade")
    session.findById("wnd[0]").sendVKey(3)  # Voltar
    return status_sistema, status_usuario, corte, relig, posicao_rede, profundidade
