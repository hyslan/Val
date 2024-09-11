"""Module to process the options chosen by the user in the main menu."""

from __future__ import annotations

import logging
import typing

import numpy as np
import win32com.client
from rich.console import Console

from python.src.core import val
from python.src.desvalorador import desvalorador
from python.src.etl import extract_from_sql, pendentes_csv, pendentes_excel
from python.src.osn3 import pertencedor
from python.src.retrabalhador import retrabalho
from python.src.sql_view import Sql

if typing.TYPE_CHECKING:
    import argparse

    import rich.console
    import win32com.client

console = Console()
logger = logging.getLogger(__name__)


# ruff: noqa: C901
def process_valoration(
    options: str,
    args: argparse.Namespace,
    session: win32com.client.CDispatch,
    token: str,
    console: rich.console.Console,
) -> None:
    """Função para processar as opções escolhidas pelo usuário no menu principal.

    Args:
    ----
        options (str): opção escolhida pelo usuário.
        args (argparse.Namespace): args passados no cli
        session (win32com.client.CDispatch): Sessão do SAPGUI
        token (str): SSO token
        console (rich.console.Console): Terminal Highlitghing

    """
    try:
        match options:
            case "1":
                desvalorador(args.contrato, session)

            case "2":
                retrabalho(args.contrato, session)

            case "3":
                pertencedor(args.contrato, session, args.session)

            case "4":
                pendentes_list: np.ndarray = extract_from_sql(args.contrato)
                val(pendentes_list, session, args.contrato, args.revalorar, token, args.session)

            case "5":
                tses_existentes = Sql("", "")
                console.print("\n", tses_existentes.show_tses(), style="italic blue", justify="full")
                tse_expec = input("- Val: Digite as TSE separadas por vírgula, por favor.\n")
                lista_tse = tse_expec.split(", ")
                pendentes = Sql(ordem="", cod_tse=lista_tse)
                pendentes_array: np.ndarray = pendentes.tse_escolhida(args.contrato)
                val(pendentes_array, session, args.contrato, args.revalorar, token, args.session)

            case "6":
                tses_existentes = Sql("", "")
                console.print("\n", tses_existentes.show_tses(), style="italic blue", justify="full")
                tse_expec = input("- Val: Digite a TSE expecífica, por favor.\n")
                pendentes = Sql(ordem="", cod_tse=tse_expec)
                pendentes_array: np.ndarray = pendentes.tse_expecifica(args.contrato)
                val(pendentes_array, session, args.contrato, args.revalorar, token, args.session)

            case "7":
                ordem_expec = input(
                    "- Val: Digite o Nº da Ordem, por favor.\n",
                )
                mun = input("Digite o Nº do Município.\n")
                pendentes_array: np.ndarray = np.array([[ordem_expec, mun]])
                val(
                    pendentes_array,
                    session,
                    args.contrato,
                    args.revalorar,
                    token,
                    args.session,
                )

            case "8":
                ask = input("é csv?")
                planilha = pendentes_csv() if ask == "s" else pendentes_excel()
                val(planilha, session, args.contrato, args.revalorar, token, args.session)

            case "9":
                pendentes = Sql("", "")
                pendentes_array = pendentes.familia(args.family, args.contrato)
                val(
                    pendentes_array,
                    session,
                    args.contrato,
                    args.revalorar,
                    token,
                    args.session,
                )

            case "10":
                pendentes = Sql("", "")
                pendentes_array = pendentes.desobstrucao()
                val(
                    pendentes_array,
                    session,
                    args.contrato,
                    args.revalorar,
                    token,
                    args.session,
                )

    except (TypeError, ValueError) as erro:
        console.print(f":fearful: Erro Main: {erro}", style="bold red")
        console.print_exception()
        logger.exception("Erro Main ", exc_info=erro)
        return
