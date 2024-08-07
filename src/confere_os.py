# ConfereOS.py
"""Módulo de check-up no SAP"""
import threading
import pythoncom
import win32com.client as win32
import pywintypes
from rich.console import Console
from src import sap
from src.lista_reposicao import dict_reposicao

# Adicionando um Lock
lock = threading.Lock()
console = Console()


def consulta_os(n_os, session, contrato, n_con):
    """
    Função para consultar ORDEM na transação ZSBPM020.

    Args:
        n_os (str): Número da ordem de serviço.
        session (win32com.client.CDispatch): Objeto de sessão SAP.
        contrato (str): Contrato relacionado à ordem de serviço.

    Returns:
        tuple: Uma tupla contendo as seguintes informações:
            - status_sistema (str): Status do sistema.
            - status_usuario (str): Status do usuário.
            - corte (str): Local de corte.
            - relig (str): Local de religação.
            - posicao_rede (str): Posição da rede.
            - profundidade (str): Profundidade.
            - hidro (str): Hidrômetro instalado.
            - operacao (str): Operação.
            - diametro_ramal (str): Diâmetro do ramal.
            - diametro_rede (str): Diâmetro da rede.
            - principal_tse (str): Código da principal tse.
    """
    status_sistema = None
    status_usuario = None
    posicao_rede = None
    profundidade = None
    operacao = None
    diametro_ramal = None
    diametro_rede = None
    hidro = None
    relig = None
    corte = None
    principal_tse = None

    def zsbpm020(session_id):
        """Transact 020"""
        nonlocal n_os
        nonlocal contrato

        # Seção Crítica - uso do Lock
        with lock:
            try:
                # pylint: disable=E1101
                pythoncom.CoInitialize()
                # pylint: disable=E1101
                gui = win32.Dispatch(
                    pythoncom.CoGetInterfaceAndReleaseStream(
                        session_id, pythoncom.IID_IDispatch)
                )
                if gui is None:
                    console.print("Não foi possível obter a sessão SAP.")
                    return
                gui.StartTransaction("ZSBPM020")
                campo_os = gui.findById("wnd[0]/usr/ctxtS_AUFNR-LOW")
                campo_os.Text = n_os
                gui.findById("wnd[0]/usr/txtS_CONTR-LOW").text = contrato
                # gui.findById("wnd[0]/usr/txtS_UN_ADM-LOW").text =
                gui.findById("wnd[0]").sendVKey(8)
            except (pywintypes.com_error, AttributeError) as transaction_error:
                console.print(
                    f"Erro durante a thread consulta OS: {transaction_error}")

            finally:
                pythoncom.CoUninitialize()

    try:
        # pylint: disable=E1101
        pythoncom.CoInitialize()
        session_id = pythoncom.CoMarshalInterThreadInterfaceInStream(
            pythoncom.IID_IDispatch, session)
        # Start
        thread = threading.Thread(target=zsbpm020, kwargs={
            'session_id': session_id})
        thread.start()
        # Aguarde a thread concluir
        thread.join(timeout=300)
        if thread.is_alive():
            print("SAP demorando mais que o esperado, encerrando.")
            sap.fechar_conexao(n_con)

        consulta = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        status_sistema = consulta.GetCellValue(0, "STTXT")
        status_usuario = consulta.GetCellValue(0, "USTXT")
        # Contagem Grid
        num_etapas_linhas = consulta.RowCount

        for n, cod_tse in enumerate(range(0, num_etapas_linhas)):
            cod_tse = consulta.GetCellValue(n, "ZZTSE")
            if cod_tse not in (
                    dict_reposicao['asfalto_frio'],
                    dict_reposicao['cimentado'],
                    dict_reposicao['especial'],
                    dict_reposicao['PARARELO'],
                    dict_reposicao['SARJETA'],
                    dict_reposicao['asfalto'],
                    dict_reposicao['bloquete_inv']):
                principal_tse = cod_tse

        for n_etapa, restab in enumerate(range(0, num_etapas_linhas)):
            restab = consulta.GetCellValue(
                n_etapa, "ZZLOCAL_RELIGA")  # religação
            if restab != '':
                relig = restab

        for n_etapa, supressao in enumerate(range(0, num_etapas_linhas)):
            supressao = consulta.GetCellValue(
                n_etapa, "ZZLOCAL_CORTE")  # Supressão
            if supressao != '':
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
                    '263000',
                    '265000',
                    '266000',
                    '268000',
                    '269000',
                    '280000',
                    '284500',
                    '286000',
                    '502000',
                    '505000',
                    '506000',
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
                    '506000',
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
            if hidro_colocado:
                hidro = hidro_colocado
                break

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

    except pywintypes.com_error as erro_consulta:
        console.print(f"Erro no confere OS: {erro_consulta}")
        console.print_exception(show_locals=True)

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
        diametro_rede,
        principal_tse
    )
