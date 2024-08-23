#!/usr/bin/env python3
"""Sistema Val.

programa de valoração automática não assistida
Author: Hyslan Silva Cruz.
"""

# main.py
# Bibliotecas
from __future__ import annotations

import argparse
import datetime
import os
import time
import typing

import numpy as np
import pywintypes
from dotenv import load_dotenv
from rich.console import Console

from python.src import sap, sql_view
from python.src.avatar import val_avatar
from python.src.core import val
from python.src.desvalorador import desvalorador
from python.src.etl import extract_from_sql, pendentes_csv, pendentes_excel
from python.src.face_the_gandalf import you_cant_pass
from python.src.log_config import setup_logging
from python.src.osn3 import pertencedor
from python.src.retrabalhador import retrabalho
from python.src.sapador import down_sap

if typing.TYPE_CHECKING:
    import rich.console
    import win32com.client


def main(args=None) -> None:
    """Sistema principal da Val e inicializador do programa."""
    setup_logging()
    # Argumentos
    parser: argparse.ArgumentParser = argparse.ArgumentParser(
        prog="Sistema Val",
        description="""Sistema de valoração automática não assistida em lote.
        Selecione uma opção de valoração e o contrato a ser utilizado. Caso deseje, escolha uma família específica.
        Por padrão, o programa irá valoração todas as ordens pendentes das famílias:
        CAVALETE, HIDROMETRO, POCO, RAMAL AGUA, RELIGACAO, SUPRESSAO
        Para desobstrução não explicite contrato ou família.
        Não esqueça de colocar a senha para iniciar o programa.
        Opcional: habilite o Revalorar para reprocessar uma Ordem.
        Opções:
        1) Desvalorador
        2) Retrabalho
        3) N3 - Pertence ao serviço principal
        4) TSEs pré-selecionais como 'carteira de serviços'
        5) TSEs específicas digitadas pelo usuário
        6) TSE específica digitada pelo usuário
        7) Ordem específica digitada pelo usuário
        8) Pendentes em CSV ou Excel
        9) Por Família
        10) Desobstrução NORTESUL""",
        epilog="Author: Hyslan Silva Cruz <hyslansilva@gmail.com>",
    )
    parser.add_argument(
        "-s",
        "--session",  # It's a connection not Session
        type=int,
        default=0,
        help="Número da sessão do SAP a ser utilizada.",
    )
    parser.add_argument(
        "-o",
        "--option",
        type=str,
        help="Escolha uma Opção de valoração.",
        choices=[str(i) for i in range(1, 11)],
    )
    parser.add_argument("-c", "--contrato", default=None, type=str, help="Escolha o Contrato a ser utilizado.")
    parser.add_argument("-f", "--family", default=None, type=str, nargs="+", help="Escolha a Família a ser utilizada.")
    parser.add_argument("-p", "--password", type=str, help="Digite a senha para iniciar o programa.")
    parser.add_argument("-r", "--revalorar", default=False, type=bool, help="Revalorar uma Ordem.")

    if args is None:
        args: argparse.Namespace = parser.parse_args()

    options: str = args.option
    hora_parada: datetime.time = datetime.time(21, 50)  # Ponto de parada às 21h50min
    console: rich.console.Console = Console()
    # Avatar.
    val_avatar()

    while True:
        console.print("\n[bold blue underline]Sistema Val[/bold blue underline] :smiley:", justify="full")
        # Obtém a hora atual
        hora_atual: datetime.time = datetime.datetime.now().time()
        hora: int = hora_atual.hour
        saudacao: str
        if hora < 12:
            saudacao = "Bom dia!"
        elif hora < 18:
            saudacao = "Boa tarde!"
        else:
            saudacao = "Boa noite!"
        console.print(f"- Val: {saudacao} :star:\n- Val: Como vai você hoje?")
        console.print(f"- Val: Hora atual: {hora_atual.strftime('%H:%M:%S')} :alarm_clock:")
        load_dotenv()

        if args.password != os.environ["PWD"]:
            console.print("Senha incorreta!\n Você não vai passar!. :mage:", style="bold")
            you_cant_pass("video")
            return

        # * Conexão ao SAP
        try:
            console.print("[i cyan] Conectando ao SAP GUI e obtendo token de acesso...")
            token = down_sap()
            session: win32com.client.CDispatch = sap.choose_connection(args.session)
        except pywintypes.com_error:
            console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
            console.print("[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
            token = down_sap()
            time.sleep(10)
            session: win32com.client.CDispatch = sap.choose_connection(args.session)

        try:
            match options:
                case "1":
                    desvalorador(args.contrato, session)

                case "2":
                    retrabalho(args.contrato, session)

                case "3":
                    pertencedor(args.contrato, session)

                case "4":
                    pendentes_list: np.ndarray = extract_from_sql(args.contrato)
                    val(pendentes_list, session, args.contrato, args.revalorar, token, args.session)

                case "5":
                    tses_existentes = sql_view.Sql("", "")
                    console.print("\n", tses_existentes.show_tses(), style="italic blue", justify="full")
                    tse_expec = input("- Val: Digite as TSE separadas por vírgula, por favor.\n")
                    lista_tse = tse_expec.split(", ")
                    pendentes = sql_view.Sql(ordem="", cod_tse=lista_tse)
                    pendentes_array: np.ndarray = pendentes.tse_escolhida(args.contrato)
                    val(pendentes_array, session, args.contrato, args.revalorar, token, args.session)

                case "6":
                    tses_existentes = sql_view.Sql("", "")
                    console.print("\n", tses_existentes.show_tses(), style="italic blue", justify="full")
                    tse_expec = input("- Val: Digite a TSE expecífica, por favor.\n")
                    pendentes = sql_view.Sql(ordem="", cod_tse=tse_expec)
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
                    pendentes = sql_view.Sql("", "")
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
                    pendentes = sql_view.Sql("", "")
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
            return

        # Loop de Parada
        if hora_atual >= hora_parada:
            console.print("A Val foi descansar.")
            console.print("- Val: até amanhã.")
            return

        # Encerramento
        console.print("- Val: Valoração finalizada, encerrando. :star:")
        return


# -Main------------------------------------------------------------------
if __name__ == "__main__":
    main()
# -End-------------------------------------------------------------------
