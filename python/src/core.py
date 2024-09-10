# core.py
"""Coração da Val."""

import contextlib
import logging
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
from python.src.pagador import precificador
from python.src.salvacao import salvar
from python.src.temporizador import cronometro_val
from python.src.transact_zsbmm216 import Transacao
from python.src.wms.consulta_estoque import estoque

from .log_decorator import log_execution

# Global Class print highlighting
console: rich.console.Console = Console()
logger = logging.getLogger(__name__)
# * Variáveis de Status da Ordem
valorada: str = "EXEC VALO"
fechada: str = "LIB"


def rollback(n: int, token: str) -> win32com.client.CDispatch:
    """Tenta começar uma nova conexão com o SAPGUI.

    Args:
    ----
        n (int): número da conexão
        token (str): Acesso SSO do usuário

    Returns:
    -------
        win32com.client.CDispatch: _description_

    """
    with contextlib.suppress(pywintypes.com_error):
        sap.fechar_conexao(n)

    sap.get_connection(token)
    time.sleep(10)
    session: win32com.client.CDispatch = sap.choose_connection(n)
    return session


@log_execution
def estoque_virtual(contrato: str, n_con: int) -> DataFrame | None:
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
    # NORTE SUL DESOBSTRUÇÃO
    if contrato not in ("4600043760", "4600046036", "4600045267", "4600043654"):
        new_session: win32com.client.CDispatch = sap.create_session(n_con)
        estoque_hj: DataFrame | None = estoque(new_session, contrato, n_con)

    if len(estoque_hj) == 0:
        estoque_hj = None


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
    """Verifica quem valorou.

    Args:
    ----
        session (win32com.client.CDispatch): Objeto COM SAPGUI x86 da sessão
        sessions (win32com.client.CDispatch): qtd de sessões Objeto COM SAPGUI x86
        ordem (str): Número da Ordem
        contrato (str): Contrato utilizado
        cod_mun (str): Código do município
        principal_tse (str): Serviço Pai
        start_time (float): Cronometro inicial da iteração
        n_con (int): número da conexão

    Returns:
    -------
        str | None: Data da valoração

    """
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
    f_total = float(total.replace(".", "").replace(",", ".")) if total != "" else 0
    if data_valorado is not None:
        time_spent = cronometro_val(start_time, ordem)
        ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
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

    if sessions.Count != max_sessions:
        con.CloseSession(new_session.ID)

    return data_valorado


@log_execution
def inspector_materials(
    chave_rb_investimento: tuple[str, str, str, list[str], list[str]] | None,
    list_chave_rb_despesa: list[tuple[str, str, str, list[str], list[str]]] | None,
    list_chave_unitario: list[tuple[str, str, str, list[str], list[str]]] | None,
    hidro: str,
    diametro_ramal: str,
    diametro_rede: str,
    contrato: str,
    estoque_hj: DataFrame,
    posicao_rede: str,
    session: win32com.client.CDispatch,
) -> None:
    """Aciona a função de materiais.

    Args:
    ----
        chave_rb_investimento (list[str, str, str, list[str], list[str]]): Tupla com a chave de investimento
        list_chave_rb_despesa (_type_): Lista de Tupla com a chave de despesa
        list_chave_unitario (_type_): Lista de Tupla com a chave unitária1
        hidro (str): Número de Série do hidro.
        diametro_ramal (str): tamanho ramal
        diametro_rede (str): tamanho rede
        contrato (str): Número do contrato
        estoque_hj (str): Dataframe do materiais
        posicao_rede (str): Posição da rede
        session (win32com.client.CDispatch): Sessão do SAPGUI

    """
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


def consulta_status_ordem(
    ordem: str,
    session: win32com.client.CDispatch,
    contrato: str,
    n_con: int,
    start_time: float,
    cod_mun: str,
    sessions: win32com.client.CDispatch,
    revalorar: bool,
) -> tuple[str, str, str, str, str, str, str, str] | bool:
    """Consulta o status de uma ordem.

    Args:
    ----
        ordem (str): O número da ordem.
        session (win32com.client.CDispatch): A sessão do cliente.
        contrato (str): O contrato relacionado à ordem.
        n_con (int): O número do contrato.
        start_time (float): O tempo de início da consulta.
        cod_mun (str): O código do município.
        sessions (win32com.client.CDispatch): As sessões do cliente.
        revalorar (bool): Indica se a ordem deve ser revalorada.

    Returns:
    -------
        tuple[str, str, str, str, str, str, str, str] | True: Uma tupla contendo o status do sistema, o status do usuário,
        o corte, a religação, a posição da rede, a profundidade, a hidro, o diâmetro do ramal,
        o diâmetro da rede e o principal TSE.
        Retorna True se a ordem não estiver fechada e não precisar ser revalorada.

    """
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
        return True

    if revalorar is False and status_usuario == valorada:
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
        return True

    return (
        corte,
        relig,
        posicao_rede,
        profundidade,
        hidro,
        diametro_ramal,
        diametro_rede,
        principal_tse,
    )


def check_and_value_ordem(
    session: win32com.client.CDispatch,
    sessions: win32com.client.CDispatch,
    ordem: str,
    contrato: str,
    cod_mun: str,
    principal_tse: str,
    start_time: float,
    n_con: int,
    revalorar: bool,
) -> bool | win32com.client.CDispatch:
    """Checa se a ordem já foi valorada.

    Args:
    ----
        session (win32com.client.CDispatch): Sessão do SAPGUI
        sessions (win32com.client.CDispatch): Count de sessões do SAPGUI
        ordem (str): Número da Ordem.
        contrato (str): Número do Contrato.
        cod_mun (str): Código do Município
        principal_tse (str): Serviço Pai
        start_time (float): Cronometro inicial da iteração
        n_con (int): Número da Conexão
        revalorar (bool): Se True ignora o histórico da ZSBMM216 e valora de novo

    Returns:
    -------
        bool | win32com.client.CDispatch: Se True pula para a próxima iteração,
        caso contrário retorna a sessão do SAPGUI com o objeto grid tse.

    """
    try:
        session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
        )
    # Tem casos de unidade Administrativa antiga (340,344,348)
    except pywintypes.com_error:
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
        return True

    # * Check if the 'Ordem' was already valued.
    if revalorar is False:
        try:
            session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
            message_not_valued = "Não há dados para exibição."
            rodape = session.findById("wnd[0]/sbar").Text
            # Livre para valorar.
            if rodape == message_not_valued:
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS",
                ).select()
                tse = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
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
                    return True

        except pywintypes.com_error:
            console.print_exception()
            return True
        else:
            return tse
    else:
        return session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
        )


def process_precificador(
    tse: win32com.client.CDispatch,
    corte: str,
    relig: str,
    posicao_rede: str,
    profundidade: str,
    contrato: str,
    session: win32com.client.CDispatch,
    ordem: str,
    cod_mun: str,
    principal_tse: str,
    start_time: float,
) -> bool | tuple[list[str], list[str], list[str]]:
    """Processo de precificação se aplicável.

    Args:
    ----
        tse (str): Grid TSE
        corte (str): Onde foi feito a Supressão (se aplicável)
        relig (str): Onde foi feito a Religação (se aplicável)
        posicao_rede (str): Posição da Rede
        profundidade (str): Profundidade da rede
        contrato (str): Número do Contrato
        session (win32com.client.CDispatch): Sessão do SAPGUI
        ordem (str): Número da Ordem.
        cod_mun (str): Código do Município
        principal_tse (str): Serviço Pai da Ordem
        start_time (float): Cronômetro inicial da iteração

    Returns:
    -------
        True | tuple[list[str], list[str], list[str]]: Chaves com a descrição do Serviço e sua classe.

    """
    (
        tse_proibida,
        list_chave_rb_despesa,
        list_chave_unitario,
        chave_rb_investimento,
        _,  # Skipping 'chave_unitario'
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
        return True

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
        return True

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
        return True
    return (
        list_chave_rb_despesa,
        list_chave_unitario,
        chave_rb_investimento,
    )


def looping(
    pendentes_array: np.ndarray,
    session: win32com.client.CDispatch,
    sessions: win32com.client.CDispatch,
    contrato: str | None,
    revalorar: bool,
    token: str,
    n_con: int,
    transacao: Transacao,
    estoque_hj: DataFrame | None,
) -> None:
    """Motor do Looping para NorteSul e Globais.

    Args:
    ----
        pendentes_array (np.ndarray): Array das ordens, cod_municipio e empresa (se for NORTESUL)
        session (win32com.client.CDispatch): Sessão do SAPGUI
        sessions (win32com.client.CDispatch): Count da Sessão do SAPGUI
        contrato (str | None): Número do Contrato.
        revalorar (bool): Se True ignora o EXEC VALO da transação ZSBPM020
        token (str): SSO de acesso.
        n_con (int): Número da Conexão
        transacao (Transacao): Classe do ZSBMM216
        estoque_hj (DataFrame | None): Dataframe do estoque virtual

    Returns:
    -------
        _type_: None

    """
    qtd_ordem: int = 0  # Contador de ordens pagas.

    def count_done(qtd_ordem: int) -> None:
        """Conta as ordens valoradas."""
        console.print(
            Panel.fit(f"Quantidade de ordens valoradas: {qtd_ordem}."),
            style="italic yellow",
        )

    def process_order(
        ordem: str,
        cod_mun: str,
        empresa: str,
        qtd_ordem: int,
        session: win32com.client.CDispatch,
        sessions: win32com.client.CDispatch,
        revalorar: bool,
        token: str,
        n_con: int,
    ) -> tuple[int, win32com.client.CDispatch]:
        """Processa a ordem de serviço em etapas.

        Args:
        ----
            ordem (str): Número da Ordem
            cod_mun (str): Código do Município
            empresa (str): Número do contrato (contrato)
            qtd_ordem (int): qtde de ordens valoradas
            session (win32com.client.CDispatch): Sessão do SAPGUI
            sessions (win32com.client.CDispatch): Sessões do SAPGUI
            revalorar (bool): Se True ignora o EXEC VALO da transação ZSBPM020
            token (str): SSO de acesso
            n_con (int): Número da Conexão

        Returns:
        -------
            tuple[int, win32com.client.CDispatch]: Qtde de ordens valoradas e sessão do SAPGUI (COM x86)

        """
        try:
            start_time = time.time()  # Contador de tempo para valorar.
            console.print(f"[b]Ordem atual: {ordem}")
            # Função consulta de Ordem.
            result = consulta_status_ordem(
                ordem,
                session,
                empresa,
                n_con,
                start_time,
                cod_mun,
                sessions,
                revalorar,
            )
            if result is True:
                return qtd_ordem, session

            if isinstance(result, tuple):
                (
                    corte,
                    relig,
                    posicao_rede,
                    profundidade,
                    hidro,
                    diametro_ramal,
                    diametro_rede,
                    principal_tse,
                ) = result

            # * Go To ZSBMM216 Transaction
            transacao.contrato = empresa if contrato is None else contrato
            transacao.municipio = cod_mun
            transacao.run_transacao(ordem)

            # * Processo de Serviços Executados
            check = check_and_value_ordem(
                session,
                sessions,
                ordem,
                empresa,
                cod_mun,
                principal_tse,
                start_time,
                n_con,
                revalorar,
            )
            if check is True:
                return qtd_ordem, session

            if isinstance(check, win32com.client.CDispatch):
                tse: win32com.client.CDispatch = check

                # * TSE e Aba Itens de preço
                result_p = process_precificador(
                    tse,
                    corte,
                    relig,
                    posicao_rede,
                    profundidade,
                    empresa,
                    session,
                    ordem,
                    cod_mun,
                    principal_tse,
                    start_time,
                )
                if result_p is True:
                    return qtd_ordem, session

                if isinstance(result_p, tuple):
                    (
                        list_chave_rb_despesa,
                        list_chave_unitario,
                        chave_rb_investimento,
                    ) = result_p

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

            # * Salvar Ordem
            qtd_ordem = salvar(
                ordem,
                qtd_ordem,
                contrato,
                session,
                principal_tse,
                cod_mun,
                start_time,
                n_con,
            )

        except pywintypes.com_error as errocritico:
            console.print_exception()
            try:
                _, descricao, _, _ = errocritico.args
                match descricao:
                    case "Falha catastrófica":
                        console.print("[bold red]SAPGUI has crashed. :fire:")
                        console.print_exception()
                        session = rollback(n_con, token)
                    case "Falha na chamada de procedimento remoto.":
                        console.print(
                            "[bold red]SAPGUI has been finished strangely. :fire:",
                        )
                        session = rollback(n_con, token)
                    case "O servidor RPC não está disponível.":
                        console.print(
                            "[bold red]SAPGUI was weirdly disconnected. :fire:",
                        )
                        session = rollback(n_con, token)
                    case _:
                        console.print(
                            "[bold red underline]Aconteceu um Erro com a Val!" + f"\n Fatal Error: {errocritico}",
                        )
                        console.print_exception()
                        logger.critical("Erro crítico na execução do looping: %s", errocritico)
            except pywintypes.com_error as e_looping:
                logger.critical(
                    "Erro crítico na execução do looping: %s",
                    e_looping,
                )
                return qtd_ordem, session
        else:
            return qtd_ordem, session

        return qtd_ordem, session

    if contrato is None:
        # NORTESUL
        for ordem, cod_mun, empresa in tqdm(pendentes_array, ncols=100):
            qtd_ordem, session = process_order(
                ordem,
                cod_mun,
                empresa,
                qtd_ordem,
                session,
                sessions,
                revalorar,
                token,
                n_con,
            )
            count_done(qtd_ordem)
    else:
        # GLOBAIS
        for ordem, cod_mun in tqdm(pendentes_array, ncols=100):
            qtd_ordem, session = process_order(
                ordem,
                cod_mun,
                contrato,
                qtd_ordem,
                session,
                sessions,
                revalorar,
                token,
                n_con,
            )
            count_done(qtd_ordem)


@log_execution
def val(
    pendentes_array: np.ndarray,
    session: win32com.client.CDispatch,
    contrato: str | None,
    revalorar: bool,
    token: str,
    n_con: int,
) -> None:
    """Sistema Val.

    Neste sistema, a Val irá valorar as ordens de serviço do SAP.
    Núcleo motor para instanciar o looping e seus ramos de cada iteração.
    """
    if isinstance(contrato, str):
        transacao: Transacao = Transacao(contrato, "100", session)

    limite_execucoes = len(pendentes_array)
    # * In case of null Df.
    if limite_execucoes == 0:
        msg = "Nenhuma ordem para ser valorada."
        raise ValueError(msg)

    try:
        sessions: win32com.client.CDispatch = sap.listar_sessoes(n_con)
    except pywintypes.com_error:
        return

    if contrato is not None:
        estoque_hj: DataFrame | None = estoque_virtual(contrato, n_con)
    else:
        estoque_hj = None

    with console.status("[bold blue]Trabalhando..."):
        looping(
            pendentes_array,
            session,
            sessions,
            contrato,
            revalorar,
            token,
            n_con,
            transacao,
            estoque_hj,
        )

    return
