# ConfereOS.py
'''Módulo de check-up no SAP'''
from sap_connection import connect_to_sap


def consulta_os(n_os):
    '''Função para consultar ORDEM na transação ZSBPM020.'''
    diametro_ramal = None
    diametro_rede = None
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
    corte = consulta.GetCellValue(0, "ZZLOCAL_CORTE")  # Supressão
    relig = consulta.GetCellValue(0, "ZZLOCAL_RELIGA")  # religação
    posicao_rede = consulta.GetCellValue(
        0, "ZZPOSICAO_REDE")  # Posição da Rede
    profundidade = consulta.GetCellValue(0, "ZZPROFUNDIDADE")
    num_etapas_linhas = consulta.RowCount
    for n_etapa, operacao in enumerate(range(0, num_etapas_linhas)):
        operacao = consulta.GetCellValue(n_etapa, "VORNR")
    for n_etapa, hidro_colocado in enumerate(range(0, num_etapas_linhas)):
        hidro_colocado = consulta.GetCellValue(
            n_etapa, "ZZHIDROMETRO_INSTALADO")
        if hidro_colocado is not None:
            hidro = hidro_colocado
    for n_etapa in range(0, num_etapas_linhas):
        sap_diametro_ramal = consulta.GetCellValue(
            n_etapa, "ZZDIAMETRO_RAMAL"
        )
        if sap_diametro_ramal in ('100', '150', '200', '250', '300'):
            diametro_ramal = sap_diametro_ramal
            break
    for n_etapa in range(0, num_etapas_linhas):
        sap_diametro_rede = consulta.GetCellValue(
            n_etapa, "ZZDIAMETRO_REDE"
        )
        if sap_diametro_rede in ('100', '150', '200', '250', '300'):
            diametro_rede = sap_diametro_rede
            break
    session.findById("wnd[0]").sendVKey(3)  # Voltar

    return (
        status_sistema,
        status_usuario,
        corte,
        relig,
        posicao_rede,
        profundidade,
        hidro,
        operacao,
        diametro_ramal,
        diametro_rede
    )
