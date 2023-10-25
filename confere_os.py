# ConfereOS.py
'''Módulo de check-up no SAP'''
import threading
from sap_connection import connect_to_sap
from sap import encerrar_sap

# Adicionando um Lock
lock = threading.Lock()


def consulta_os(n_os):
    '''Função para consultar ORDEM na transação ZSBPM020.'''
    diametro_ramal = None
    diametro_rede = None

    def zsbpm020():
        '''Transact 020'''
        nonlocal n_os

        # Seção Crítica - uso do Lock
        with lock:
            session = connect_to_sap()
            session.StartTransaction("ZSBPM020")
            campo_os = session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW")
            campo_os.Text = n_os
            session.findById("wnd[0]/usr/txtS_CONTR-LOW").text = "4600041302"
            session.findById("wnd[0]/usr/txtS_UN_ADM-LOW").text = "344"
            session.findById("wnd[0]").sendVKey(8)

    # Start thread save.
    thread = threading.Thread(target=zsbpm020)
    thread.start()
    # Timeout 5min
    thread.join(timeout=300)
    if thread.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()

    session = connect_to_sap()
    consulta = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    status_sistema = consulta.GetCellValue(0, "STTXT")
    status_usuario = consulta.GetCellValue(0, "USTXT")
    # Contagem Grid
    num_etapas_linhas = consulta.RowCount

    for n_etapa, restab in enumerate(range(0, num_etapas_linhas)):
        restab = consulta.GetCellValue(n_etapa, "ZZLOCAL_RELIGA")  # religação
        if restab is not None:
            relig = restab

    for n_etapa, supressao in enumerate(range(0, num_etapas_linhas)):
        supressao = consulta.GetCellValue(
            n_etapa, "ZZLOCAL_CORTE")  # Supressão
        if supressao is not None:
            corte = supressao

    for n_etapa in range(0, num_etapas_linhas):
        tipo_tse = consulta.GetCellValue(
            n_etapa, "ZZTSE"
        )
        if tipo_tse in (
            '253000',
            '254000',
            '255000',
            '262000',
            '265000',
            '266000',
            '268000',
            '269000',
            '280000',
            '284500',
            '286000',
            '502000',
            '505000',
            '508000',
            '561000',
            '565000',
            '569000',
            '581000',
            '539000',
            '539000',
            '585000',
        ):
            # Posição da Rede
            posicao_rede = consulta.GetCellValue(n_etapa, "ZZPOSICAO_REDE")
            break
        else:
            posicao_rede = None

    for n_etapa in range(0, num_etapas_linhas):
        tipo_tse = consulta.GetCellValue(
            n_etapa, "ZZTSE"
        )
        if tipo_tse in (
            '502000',
            '505000',
            '508000',
            '561000',
            '565000',
            '569000',
            '581000',
            '539000',
            '539000',
            '585000',
        ):
            # Profundidade
            profundidade = consulta.GetCellValue(n_etapa, "ZZPROFUNDIDADE")
            break
        else:
            profundidade = None

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
        if sap_diametro_ramal in ('DN_100', 'DN_150', 'DN_200', 'DN_250', 'DN_300'):
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
