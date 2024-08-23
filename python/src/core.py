# core.py
"""Coração da Val."""

import sys
import time

import numpy as np
import pywintypes
import rich.console
import win32com.client
from pandas import DataFrame
from rich.console import Console
from rich.panel import Panel
from tqdm import tqdm

from python.src import sap, sql_view
from python.src.almoxarifado import materiais
from python.src.confere_os import consulta_os
from python.src.nazare_bugou import oxe
from python.src.pagador import precificador
from python.src.salvacao import salvar
from python.src.temporizador import cronometro_val
from python.src.transact_zsbmm216 import Transacao
from python.src.wms.consulta_estoque import estoque

from .log_decorator import log_execution

# Global Class print highlighting
console: rich.console.Console = Console()


def rollback(n: int, token) -> win32com.client.CDispatch:
    try:
        sap.fechar_conexao(n)
    except pywintypes.com_error:
        print("Conexão já Encerrada ou não ativa.")

    sap.get_connection(token)
    time.sleep(10)
    session: win32com.client.CDispatch = sap.choose_connection(n)
    print("Sessão Recuperada.")
    return session


@log_execution
def estoque_virtual(contrato, n_con) -> DataFrame | None:
    """Get the virtual stock in MBLB transaction using contrato's number.

    Args:
    ----
        contrato (str): _description_
        n_con (win32com.client.CDispatch): Connection number.
        sessions (win32com.client.CDispatch): Len of Sessions active.

    Raises:
    ------
        Exception: Error Message.

    Returns:
    -------
        DataFrame: Pandas Dataframe with all material allocated to the 'contrato'

    """
    try:
        # NORTE SUL DESOBSTRUÇÃO
        if contrato not in ("4600043760", "4600046036", "4600045267", "4600043654"):
            new_session: win32com.client.CDispatch = sap.create_session(n_con)
            estoque_hj: DataFrame = estoque(new_session, contrato, n_con)

        return estoque_hj

    except Exception as e_estoque_v:
        console.print(f"[b] Erro ao obter o estoque virtual: {e_estoque_v}")
        console.print_exception()
        # raise Exception("Extração do Estoque Virtual Falhou!")


@log_execution
def valorator_user(
    session: win32com.client.CDispatch,
    sessions: win32com.client.CDispatch,
    ordem: str,
    contrato: str,
    cod_mun: str,
    principal_tse: str,
    start_time: float,
    n_con: int,
) -> str | None:
    data_valorado = None
    max_sessions = 6
    if sessions.Count != max_sessions:
        new_session: win32com.client.CDispatch = sap.create_session(n_con)
    else:
        new_session = session

    con = sap.connection_object(n_con)

    transaction_check: Transacao = Transacao(contrato, cod_mun, new_session)
    transaction_check.run_transacao(ordem, tipo="consulta")

    new_session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
    grid_historico = new_session.findById(
        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA/ssubSUB_TAB:" + "ZSBMM_VALORACAO_NAPI:9040/cntlCC_AJUSTES/shellcont/shell",
    )
    data_valorado = grid_historico.GetCellValue(0, "DATA")
    matricula = grid_historico.GetCellValue(0, "MODIFICADO")
    total = new_session.findById("wnd[0]/usr/txtGS_HEADER-VAL_ATUAL").Text
    if total != "":
        f_total = float(total.replace(".", "").replace(",", "."))
    else:
        f_total = 0
    if data_valorado is not None:
        time_spent = cronometro_val(start_time, ordem)
        print(f"OS: {ordem} já valorada.")
        print(f"Data: {data_valorado}")
        print(f"Matrícula: {matricula}")
        print("Valor Medido: ", f_total)
        ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
        try:
            ja_valorado.valorada(
                valorado="SIM",
                contrato=contrato,
                municipio=cod_mun,
                status="VALORADA",
                obs="",
                data_valoracao=data_valorado,
                matricula=matricula,
                valor_medido=f_total,
                tempo_gasto=time_spent,
            )
            ja_valorado.clean_duplicates()
        except Exception as e_valorado:
            console.print(f"[i yellow]Erro no SQL de valorator_user: {e_valorado}")

    if sessions.Count != 6:
        con.CloseSession(new_session.ID)

    return data_valorado


@log_execution
def inspector_materials(
    chave_rb_investimento,
    list_chave_rb_despesa,
    list_chave_unitario,
    hidro,
    diametro_ramal,
    diametro_rede,
    contrato,
    estoque_hj,
    posicao_rede,
    session,
) -> None:
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
            session,
        )
    # RB - Despesa
    # NORTESUL
    if list_chave_rb_despesa and contrato not in (
        "4600043760",
        "4600046036",
        "4600045267",
    ):
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
                session,
            )
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
                session,
            )


@log_execution
def val(
    pendentes_array: np.ndarray,
    session: win32com.client.CDispatch,
    contrato: str | None,
    revalorar: bool,
    token: str,
    n_con: int,
) -> None:
    """Sistema Val."""
    transacao: Transacao = Transacao(contrato, "100", session)
    limite_execucoes = len(pendentes_array)
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    # * In case of null Df.
    if limite_execucoes == 0:
        print("Nenhuma Ordem para Valorar.")
        return True

    try:
        sessions: win32com.client.CDispatch = sap.listar_sessoes(n_con)
    except pywintypes.com_error:
        return None

    if contrato is not None:
        estoque_hj: DataFrame | None = estoque_virtual(contrato, n_con)

    with console.status("[bold blue]Trabalhando..."):
        # * Variáveis de Status da Ordem
        valorada: str = "EXEC VALO" or "NEXE VALO"
        fechada: str = "LIB"
        qtd_ordem: int = 0  # Contador de ordens pagas.
        # * For NORTESUL
        if contrato is None:
            for ordem, cod_mun, empresa in tqdm(pendentes_array, ncols=100):
                try:
                    start_time = time.time()  # Contador de tempo para valorar.
                    console.print(f"[b]Ordem atual: {ordem}")
                    print("Verificando Status da Ordem.")
                    # Função consulta de Ordem.
                    print("Iniciando Consulta.")
                    (
                        status_sistema,
                        status_usuario,
                        corte,
                        relig,
                        posicao_rede,
                        profundidade,
                        hidro,
                        _,  # Skipping 'operacao'
                        diametro_ramal,
                        diametro_rede,
                        principal_tse,
                    ) = consulta_os(ordem, session, empresa, n_con)
                    # * Consulta Status da Ordem
                    if status_sistema != fechada:
                        print(f"OS: {ordem} aberta.")
                        time_spent = cronometro_val(start_time, ordem)
                        ja_valorado = sql_view.Sql(ordem, principal_tse)
                        ja_valorado.valorada(
                            valorado="NÃO",
                            contrato=empresa,
                            municipio=cod_mun,
                            status="ABERTA",
                            obs="",
                            data_valoracao=None,
                            matricula="117615",
                            valor_medido=0,
                            tempo_gasto=time_spent,
                        )
                        continue

                    if revalorar is False and status_usuario == valorada:
                        print(f"OS: {ordem} já valorada.")
                        valorator_user(
                            session,
                            sessions,
                            ordem,
                            empresa,
                            cod_mun,
                            principal_tse,
                            start_time,
                            n_con,
                        )
                        continue

                    # * Go To ZSBMM216 Transaction
                    transacao.contrato = empresa
                    transacao.municipio = cod_mun
                    transacao.run_transacao(ordem)

                    console.print(
                        "Processo de Serviços Executados",
                        style="bold red underline",
                        justify="center",
                    )
                    try:
                        tse = session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                            + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                        )
                    except pywintypes.com_error:
                        print(f"Ordem: {ordem} em medição definitiva.")
                        time_spent = cronometro_val(start_time, ordem)
                        ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
                        ja_valorado.valorada(
                            valorado="SIM",
                            contrato=empresa,
                            municipio=cod_mun,
                            # Open zsbmm216 and get the date of the last valuation.
                            status="DEFINITIVA",
                            obs="",
                            data_valoracao=None,
                            matricula="",
                            valor_medido=0,
                            tempo_gasto=time_spent,
                        )
                        ja_valorado.clean_duplicates()
                        continue

                    # * Check if the 'Ordem' was already valued.
                    try:
                        session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
                        MESSAGE_NOT_VALUED = "Não há dados para exibição."
                        rodape = session.findById("wnd[0]/sbar").Text
                        if rodape == MESSAGE_NOT_VALUED:
                            print(f"OS: {ordem} não foi valorada.")
                            print("OS Livre para valorar.")
                            session.findById(
                                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS",
                            ).select()
                            tse = session.findById(
                                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                                + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                            )
                        else:
                            data_valorado = valorator_user(
                                session,
                                sessions,
                                ordem,
                                empresa,
                                cod_mun,
                                principal_tse,
                                start_time,
                                n_con,
                            )
                            if data_valorado is not None:
                                continue

                    except pywintypes.com_error:
                        console.print_exception()
                        sys.exit()

                    # * TSE e Aba Itens de preço
                    (
                        tse_proibida,
                        list_chave_rb_despesa,
                        list_chave_unitario,
                        chave_rb_investimento,
                        chave_unitario,  # Skipping 'chave_unitario'
                        ligacao_errada,
                        profundidade_errada,
                    ) = precificador(
                        tse,
                        corte,
                        relig,
                        posicao_rede,
                        profundidade,
                        empresa,
                        session,
                    )

                    # * Salvar Ordem
                    qtd_ordem, rodape = salvar(
                        ordem,
                        qtd_ordem,
                        empresa,
                        session,
                        principal_tse,
                        cod_mun,
                        start_time,
                        n_con,
                    )

                    console.print(
                        Panel.fit(f"Quantidade de ordens valoradas: {qtd_ordem}."),
                        style="italic yellow",
                    )

                except Exception as errocritico:
                    console.print_exception()
                    print(f"args do errocritico: {errocritico}")
                    try:
                        _, descricao, _, _ = errocritico.args
                        match descricao:
                            case "Falha catastrófica":
                                console.print("[bold red]SAPGUI has crashed. :fire:")
                                console.print_exception()
                                session = rollback(n_con, token)
                                continue
                            case "Falha na chamada de procedimento remoto.":
                                console.print(
                                    "[bold red]SAPGUI has been finished strangely. :fire:",
                                )
                                session = rollback(n_con, token)
                                continue
                            case "O servidor RPC não está disponível.":
                                console.print(
                                    "[bold red]SAPGUI was weirdly disconnected. :fire:",
                                )
                                session = rollback(n_con, token)
                                continue
                            case _:
                                console.print(
                                    "[bold red underline]Aconteceu um Erro com a Val!" + f"\n Fatal Error: {errocritico}",
                                )
                                console.print_exception()
                                oxe()
                                break
                    except:
                        console.print(errocritico)
            return None
        # * Loop to pay service's orders
        for ordem, cod_mun in tqdm(pendentes_array, ncols=100):
            try:
                start_time = time.time()  # Contador de tempo para valorar.
                console.print(f"[b]Ordem atual: {ordem}")
                print("Verificando Status da Ordem.")
                # Função consulta de Ordem.
                print("Iniciando Consulta.")
                (
                    status_sistema,
                    status_usuario,
                    corte,
                    relig,
                    posicao_rede,
                    profundidade,
                    hidro,
                    _,  # Skipping 'operacao'
                    diametro_ramal,
                    diametro_rede,
                    principal_tse,
                ) = consulta_os(ordem, session, contrato, n_con)
                # * Consulta Status da Ordem
                if status_sistema != fechada:
                    print(f"OS: {ordem} aberta.")
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(ordem, principal_tse)
                    ja_valorado.valorada(
                        valorado="NÃO",
                        contrato=contrato,
                        municipio=cod_mun,
                        status="ABERTA",
                        obs="",
                        data_valoracao=None,
                        matricula="117615",
                        valor_medido=0,
                        tempo_gasto=time_spent,
                    )
                    continue

                if revalorar is False:
                    if status_usuario == valorada:
                        print(f"OS: {ordem} já valorada.")
                        valorator_user(
                            session,
                            sessions,
                            ordem,
                            contrato,
                            cod_mun,
                            principal_tse,
                            start_time,
                            n_con,
                        )
                        continue

                # * Go To ZSBMM216 Transaction
                transacao.municipio = cod_mun
                transacao.run_transacao(ordem)

                console.print(
                    "Processo de Serviços Executados",
                    style="bold red underline",
                    justify="center",
                )
                try:
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                    )
                # pylint: disable=E1101
                except pywintypes.com_error:
                    print(f"Ordem: {ordem} em medição definitiva.")
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
                    ja_valorado.valorada(
                        valorado="SIM",
                        contrato=contrato,
                        municipio=cod_mun,
                        # Open zsbmm216 and get the date of the last valuation.
                        status="DEFINITIVA",
                        obs="",
                        data_valoracao=None,
                        matricula="",
                        valor_medido=0,
                        tempo_gasto=time_spent,
                    )
                    ja_valorado.clean_duplicates()
                    continue

                # * Check if the 'Ordem' was already valued.
                try:
                    session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
                    MESSAGE_NOT_VALUED = "Não há dados para exibição."
                    rodape = session.findById("wnd[0]/sbar").Text
                    if rodape == MESSAGE_NOT_VALUED:
                        print(f"OS: {ordem} não foi valorada.")
                        print("OS Livre para valorar.")
                        session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS",
                        ).select()
                        tse = session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                            + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                        )
                    else:
                        data_valorado = valorator_user(
                            session,
                            sessions,
                            ordem,
                            contrato,
                            cod_mun,
                            principal_tse,
                            start_time,
                            n_con,
                        )
                        if data_valorado is not None:
                            continue

                # pylint: disable=E1101
                except pywintypes.com_error:
                    console.print_exception()
                    exit()

                # * TSE e Aba Itens de preço
                (
                    tse_proibida,
                    list_chave_rb_despesa,
                    list_chave_unitario,
                    chave_rb_investimento,
                    chave_unitario,  # Skipping 'chave_unitario'
                    ligacao_errada,
                    profundidade_errada,
                ) = precificador(
                    tse,
                    corte,
                    relig,
                    posicao_rede,
                    profundidade,
                    contrato,
                    session,
                )
                # ! debug
                # exit()
                if ligacao_errada is True:
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
                    ja_valorado.valorada(
                        obs="Sem posição de rede.",
                        valorado="NÃO",
                        contrato=contrato,
                        municipio=cod_mun,
                        status="DISPONÍVEL",
                        data_valoracao=None,
                        matricula="117615",
                        valor_medido=0,
                        tempo_gasto=time_spent,
                    )
                    ja_valorado.clean_duplicates()
                    continue

                if profundidade_errada is True:
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
                    ja_valorado.valorada(
                        obs="Sem profundidade do ramal.",
                        valorado="NÃO",
                        contrato=contrato,
                        municipio=cod_mun,
                        status="DISPONÍVEL",
                        data_valoracao=None,
                        matricula="117615",
                        valor_medido=0,
                        tempo_gasto=time_spent,
                    )
                    ja_valorado.clean_duplicates()
                    continue

                # Se a TSE não estiver no escopo da Val, vai pular pra próxima OS.
                if tse_proibida is not None:
                    time_spent = cronometro_val(start_time, ordem)
                    ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
                    ja_valorado.valorada(
                        obs="Num Pode",
                        valorado="NÃO",
                        contrato=contrato,
                        municipio=cod_mun,
                        status="DISPONÍVEL",
                        data_valoracao=None,
                        matricula="117615",
                        valor_medido=0,
                        tempo_gasto=time_spent,
                    )
                    ja_valorado.clean_duplicates()
                    continue

                # * Aba Materiais
                if contrato is not None:
                    inspector_materials(
                        chave_rb_investimento,
                        list_chave_rb_despesa,
                        list_chave_unitario,
                        hidro,
                        diametro_ramal,
                        diametro_rede,
                        contrato,
                        estoque_hj,
                        posicao_rede,
                        session,
                    )
                # Fim dos materiais
                # ! debug
                # exit()
                # * Salvar Ordem
                qtd_ordem, rodape = salvar(
                    ordem,
                    qtd_ordem,
                    contrato,
                    session,
                    principal_tse,
                    cod_mun,
                    start_time,
                    n_con,
                )
                # ! debug
                # break

                console.print(
                    Panel.fit(f"Quantidade de ordens valoradas: {qtd_ordem}."),
                    style="italic yellow",
                )

            except Exception as errocritico:
                console.print_exception()
                print(f"args do errocritico: {errocritico}")
                try:
                    _, descricao, _, _ = errocritico.args
                    match descricao:
                        case "Falha catastrófica":
                            console.print("[bold red]SAPGUI has crashed. :fire:")
                            console.print_exception()
                            session = rollback(n_con, token)
                            continue
                        case "Falha na chamada de procedimento remoto.":
                            console.print(
                                "[bold red]SAPGUI has been finished strangely. :fire:",
                            )
                            session = rollback(n_con, token)
                            continue
                        case "O servidor RPC não está disponível.":
                            console.print(
                                "[bold red]SAPGUI was weirdly disconnected. :fire:",
                            )
                            session = rollback(n_con, token)
                            continue
                        case _:
                            console.print(
                                "[bold red underline]Aconteceu um Erro com a Val!" + f"\n Fatal Error: {errocritico}",
                            )
                            console.print_exception()
                            oxe()
                            break
                except:
                    console.print(errocritico)
                    # TODO (Hyslan):  needs a log
    return None
