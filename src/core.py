# core.py
"""Coração da Val."""
# pylint: disable=W0611
import logging
import datetime as dt
import time
from typing import Union
import numpy as np
import pywintypes
import rich.console
import win32com.client
from pandas import DataFrame
from tqdm import tqdm
from rich.console import Console
from rich.panel import Panel
from src.sap_connection import populate_sessions
from src.sapador import down_sap
from src import sql_view
from src import sap
from src.transact_zsbmm216 import Transacao
from src.confere_os import consulta_os
from src.pagador import precificador
from src.almoxarifado import materiais
from src.salvacao import salvar
from src.temporizador import cronometro_val
from src.wms.consulta_estoque import estoque
from src.nazare_bugou import oxe

# Global Class print highlighting
console: rich.console.Console = Console()


def rollback(n: int) -> win32com.client.CDispatch:
    try:
        sap.encerrar_sap()
    except:
        print("SAPLOGON já foi encerrado.")
    down_sap()
    populate_sessions()
    time.sleep(20)
    session: win32com.client.CDispatch = sap.escolher_sessao(n)
    print("Reiniciando programa")
    return session


def estoque_virtual(contrato, sessions, session) -> DataFrame:
    """Get the virtual stock in MBLB transaction using contrato's number.

    Args:
        contrato (str): _description_
        sessions (win32com.client.CDispatch): Len of Sessions active.
        session (win32com.client.CDispatch): Session in use.

    Raises:
        Exception: Error Message.

    Returns:
        DataFrame: Pandas Dataframe with all material allocated to the 'contrato'
    """
    try:
        # NORTE SUL DESOBSTRUÇÃO
        if not contrato == "4600043760":
            if not sessions.Count == 6:
                new_session: win32com.client.CDispatch = sap.criar_sessao(
                    sessions)
                estoque_hj: DataFrame = estoque(
                    new_session, sessions, contrato)
            else:
                estoque_hj: DataFrame = estoque(session, sessions, contrato)

        return estoque_hj

    except Exception as e_estoque_v:
        console.print(f"[b] Erro ao obter o estoque virtual: {e_estoque_v}")
        raise Exception("Extração do Estoque Virtual Falhou!")


def valorator_user(session, sessions, ordem, contrato, cod_mun, principal_tse, start_time) -> Union[str, None]:
    data_valorado = None
    if not sessions.Count == 6:
        new_session: win32com.client.CDispatch = sap.criar_sessao(
            sessions)
    else:
        new_session = session

    con = sap.listar_conexoes()

    transaction_check: Transacao = Transacao(contrato, cod_mun, new_session)
    transaction_check.run_transacao(ordem, tipo="consulta")

    new_session.findById(
        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
    grid_historico = new_session.findById(
        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA/ssubSUB_TAB:"
        + "ZSBMM_VALORACAO_NAPI:9040/cntlCC_AJUSTES/shellcont/shell")
    data_valorado = grid_historico.GetCellValue(
        0, "DATA")
    matricula = grid_historico.GetCellValue(0, "MODIFICADO")
    total = new_session.findById("wnd[0]/usr/txtGS_HEADER-VAL_ATUAL").Text
    f_total = float(total.replace(".", "").replace(",", "."))
    if data_valorado is not None:
        time_spent = cronometro_val(start_time, ordem)
        print(f"OS: {ordem} já valorada.")
        print(f"Data: {data_valorado}")
        print(f"Matrícula: {matricula}")
        print(f"Valor Medido: ", f_total)
        ja_valorado = sql_view.Sql(
            ordem=ordem, cod_tse=principal_tse)
        try:
            ja_valorado.valorada(
                valorado="SIM", contrato=contrato, municipio=cod_mun,
                status="VALORADA", obs='', data_valoracao=data_valorado,
                matricula=matricula, valor_medido=f_total, tempo_gasto=time_spent
            )
            ja_valorado.clean_duplicates()
        except Exception as e_valorado:
            print("Erro no SQL.")
            print(f"Erro em valorator_user: {e_valorado}")

    if not sessions.Count == 6:
        con.CloseSession(new_session.ID)

    return data_valorado


def inspector_materials(
        chave_rb_investimento, list_chave_rb_despesa, list_chave_unitario,
        hidro, diametro_ramal, diametro_rede, contrato, estoque_hj, posicao_rede, session) -> None:
    # RB - Investimento
    if chave_rb_investimento:
        materiais(
            hidro,
            chave_rb_investimento[1],  # Etapa atual da chave
            chave_rb_investimento,
            diametro_ramal,
            diametro_rede,
            contrato,
            estoque_hj,
            posicao_rede,
            session)
    # RB - Despesa
    if list_chave_rb_despesa and not contrato == "4600043760":
        for chave_rb_despesa in list_chave_rb_despesa:
            materiais(
                hidro,
                chave_rb_despesa[1],  # Etapa atual da chave
                chave_rb_despesa,
                diametro_ramal,
                diametro_rede,
                contrato,
                estoque_hj,
                posicao_rede,
                session)
    # Unitários
    if list_chave_unitario:
        for chave_unitario in list_chave_unitario:
            materiais(
                hidro,
                chave_unitario[1],  # Etapa atual da chave
                chave_unitario,
                diametro_ramal,
                diametro_rede,
                contrato,
                estoque_hj,
                posicao_rede,
                session)


def val(pendentes_array: np.ndarray, session, contrato: str, revalorar: bool):
    """Sistema Val."""
    transacao: Transacao = Transacao(contrato, "100", session)

    try:
        sessions: win32com.client.CDispatch = sap.listar_sessoes()
    # pylint: disable=E1101
    except pywintypes.com_error:
        return

    estoque_hj: DataFrame = estoque_virtual(contrato, sessions, session)

    limite_execucoes = len(pendentes_array)
    print(
        f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    # * In case of null Df.
    if limite_execucoes == 0:
        return None, True

    with console.status("[bold blue]Trabalhando..."):
        # * Variáveis de Status da Ordem
        valorada: str = "EXEC VALO" or "NEXE VALO"
        fechada: str = "LIB"
        qtd_ordem: int = 0  # Contador de ordens pagas.
        # Loop to pay service's orders
        for ordem, cod_mun in tqdm(pendentes_array, ncols=100):
            try:
                start_time = time.time()  # Contador de tempo para valorar.
                console.print(f"[b]Ordem atual: {ordem}")
                print("Verificando Status da Ordem.")
                # Função consulta de Ordem.
                print("Iniciando Consulta.")
                (status_sistema,
                 status_usuario,
                 corte,
                 relig,
                 posicao_rede,
                 profundidade,
                 hidro,
                 _,  # Skipping 'operacao'
                 diametro_ramal,
                 diametro_rede,
                 principal_tse
                 ) = consulta_os(ordem, session, contrato)
                # * Consulta Status da Ordem
                if not status_sistema == fechada:
                    print(f"OS: {ordem} aberta.")
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(ordem, principal_tse)
                    ja_valorado.valorada(
                        valorado="NÃO", contrato=contrato, municipio=cod_mun,
                        status="ABERTA", obs='', data_valoracao=None,
                        matricula='117615', valor_medido=0, tempo_gasto=time_spent
                    )
                    continue

                if revalorar is False:
                    if status_usuario == valorada:
                        print(f"OS: {ordem} já valorada.")
                        valorator_user(
                            session, sessions, ordem, contrato, cod_mun, principal_tse, start_time)
                        continue

                # * Go To ZSBMM216 Transaction
                transacao.municipio = cod_mun
                transacao.run_transacao(ordem)

                console.print("Processo de Serviços Executados",
                              style="bold red underline", justify="center")
                try:
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                    )
                # pylint: disable=E1101
                except pywintypes.com_error:
                    print(f"Ordem: {ordem} em medição definitiva.")
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(
                        ordem=ordem, cod_tse=principal_tse)
                    ja_valorado.valorada(
                        valorado="SIM", contrato=contrato, municipio=cod_mun,
                        # Open zsbmm216 and get the date of the last valuation.
                        status="DEFINITIVA", obs='', data_valoracao=None,
                        matricula='', valor_medido=0, tempo_gasto=time_spent
                    )
                    ja_valorado.clean_duplicates()
                    continue

                # * Check if the 'Ordem' was already valued.
                try:
                    data_valorado = valorator_user(
                        session, sessions, ordem, contrato, cod_mun, principal_tse, start_time
                    )
                    if data_valorado is not None:
                        continue

                # pylint: disable=E1101
                except pywintypes.com_error:
                    print("OS Livre para valorar.")
                    session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS").select()
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                    )

                # * TSE e Aba Itens de preço
                (
                    tse_proibida,
                    list_chave_rb_despesa,
                    list_chave_unitario,
                    chave_rb_investimento,
                    _,  # Skipping 'chave_unitario'
                    ligacao_errada,
                    profundidade_errada
                ) = precificador(tse, corte, relig,
                                 posicao_rede, profundidade, contrato, session)
                # ! debug
                # exit()
                if ligacao_errada is True:
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(
                        obs="Sem posição de rede.",
                        valorado="NÃO", contrato=contrato, municipio=cod_mun,
                        status="DISPONÍVEL", data_valoracao=None,
                        matricula='117615', valor_medido=0, tempo_gasto=time_spent)
                    ja_valorado.clean_duplicates()
                    continue

                if profundidade_errada is True:
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(
                        obs="Sem profundidade do ramal.",
                        valorado="NÃO", contrato=contrato, municipio=cod_mun,
                        status="DISPONÍVEL", data_valoracao=None,
                        matricula='117615', valor_medido=0, tempo_gasto=time_spent)
                    ja_valorado.clean_duplicates()
                    continue

                # Se a TSE não estiver no escopo da Val, vai pular pra próxima OS.
                if tse_proibida is not None:
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(
                        obs="Num Pode",
                        valorado="NÃO", contrato=contrato, municipio=cod_mun,
                        status="DISPONÍVEL", data_valoracao=None,
                        matricula='117615', valor_medido=0, tempo_gasto=time_spent)
                    ja_valorado.clean_duplicates()
                    continue

                # * Aba Materiais
                inspector_materials(chave_rb_investimento, list_chave_rb_despesa,
                                    list_chave_unitario, hidro, diametro_ramal,
                                    diametro_rede, contrato, estoque_hj, posicao_rede, session)
                # Fim dos materiais
                # ! debug
                # exit()
                # * Salvar Ordem
                qtd_ordem, rodape = salvar(
                    ordem, qtd_ordem, contrato, session, principal_tse, cod_mun, start_time)
                salvo = "Ajustes de valoração salvos com sucesso."
                if not salvo == rodape:
                    console.print(
                        f"Ordem: {ordem} não foi salva.", style="italic red")
                    console.print(f"[bold yellow]Motivo: {rodape}")
                    # TODO: Send to tb_valoradas 'NÃO' and reason why.
                    continue
                # ! debug
                    # break
                # Fim do contador de valoração.
                cronometro_val(start_time, ordem)
                console.print(
                    Panel.fit(
                        f"Quantidade de ordens valoradas: {qtd_ordem}."),
                    style="italic yellow")

            # pylint: disable=E1101
            except Exception as errocritico:
                print(f"args do errocritico: {errocritico}")
                try:
                    _, descricao, _, _ = errocritico.args
                    match descricao:
                        case 'Falha catastrófica':
                            console.print(
                                "[bold red]SAPGUI has crashed. :fire:")
                            console.print_exception()
                            # session = rollback(session_n)
                            # continue
                            break
                        case 'Falha na chamada de procedimento remoto.':
                            console.print(
                                "[bold red]SAPGUI has been finished strangely. :fire:")
                            # session = rollback(session_n)
                            # continue
                            break
                        case 'O servidor RPC não está disponível.':
                            console.print(
                                "[bold red]SAPGUI was weirdly disconnected. :fire:")
                            # session = rollback(session_n)
                            # continue
                            break
                        case _:
                            console.print(
                                "[bold red underline]Aconteceu um Erro com a Val!"
                                + f"\n Fatal Error: {errocritico}")
                            console.print_exception()
                            oxe()
                except:
                    console.print(errocritico)
                    # TODO: needs a log

        validador = True

    return ordem, validador
