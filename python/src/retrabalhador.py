"""Módulo dos retrabalho da valoração."""

from __future__ import annotations

import datetime as dt
import typing

import pywintypes
from rich.console import Console
from tqdm import tqdm

from python.src.sql_view import Sql
from python.src.transact_zsbmm216 import Transacao

if typing.TYPE_CHECKING:
    import rich
    from win32com.client import CDispatch

import pytz

console: rich.console.Console = Console()


def tag_n7(servico: CDispatch, session: CDispatch) -> str:
    """Aponta N7 ao serviço.

    Args:
    ----
        servico (CDispatch): GRID do SAP
        session (CDispatch): Sessão do SAP

    Returns:
    -------
        str: Rodapé

    """
    num_tse_linhas = servico.RowCount
    for n_tse in range(num_tse_linhas):
        servico.modifyCell(n_tse, "PAGAR", "n")
        # Retrabalho
        servico.modifyCell(n_tse, "CODIGO", "7")

    servico.pressEnter()
    session.findById("wnd[0]").sendVKey(11)
    session.findById("wnd[1]/usr/btnBUTTON_1").press()
    return session.findById("wnd[0]/sbar").Text  # Rodapé


# consulta os só precisa de status_sistema, status_usuario


def retrabalho(contrato: str, session: CDispatch) -> None:
    """Função Retrabalhador: N7."""
    salvo = "Ajustes de valoração salvos com sucesso."
    transacao = Transacao(contrato, "100", session)

    # SQL values
    today = dt.datetime.now(pytz.timezone("America/Sao_Paulo")).date()
    dt_fim = today.replace(day=1) - dt.timedelta(days=1)
    dt_inicio = dt_fim.replace(day=1)
    sql = Sql("", "")
    pendentes_array = sql.retrabalho_search(dt_inicio, dt_fim)
    limite_execucoes = len(pendentes_array)
    if limite_execucoes == 0:
        return

    with console.status("[bold blue]Trabalhando..."):
        for ordem, cod_mun, empresa in tqdm(pendentes_array, ncols=100):
            # Setando valores para a transação
            transacao.municipio = cod_mun
            transacao.contrato = empresa
            transacao.run_transacao(ordem)

            console.print("Processo de Serviços Executados", style="bold red underline", justify="center")
            try:
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                )
            except pywintypes.com_error:
                continue

            try:
                session.findById("wnd[0]").SendVkey(2)
                console.print("Desvalorando!")
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
                ok = "Valoração desfeita com sucesso."
                rodape = session.findById("wnd[0]/sbar").Text  # Rodapé
                if ok == rodape:
                    console.print(ok)
                    transacao.run_transacao(ordem)
                    servico = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                    )

                    rodape = tag_n7(servico, session)
                    if rodape == salvo:
                        console.print("[italic green]Foi Salvo com sucesso! :rocket:")
                    else:
                        console.print(f"Ordem: {ordem} não foi salva. :pouting_face:", style="italic yellow")

            except pywintypes.com_error:
                console.print(f"Ordem: {ordem} não valorada.")
                servico = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                )
                tag_n7(servico, session)

            # Clean Duplicates from tb_Valoradas
            sql.clean_duplicates()

    # Fim do retrabalhador
