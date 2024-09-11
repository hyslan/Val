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
import logging
import os
import sys
import time
import typing

import pytz
import pywintypes
from dotenv import load_dotenv
from rich.console import Console

from python.src import sap
from python.src.avatar import val_avatar
from python.src.face_the_gandalf import you_cant_pass
from python.src.log_config import setup_logging
from python.src.process_options import process_valoration
from python.src.sapador import down_sap

if typing.TYPE_CHECKING:
    import rich.console
    import win32com.client

console: rich.console.Console = Console()


def main() -> None:
    """Sistema principal da Val e inicializador do programa."""
    setup_logging()
    logger = logging.getLogger(__name__)
    console.print("Args Inseridos: ", sys.argv)
    # Argumentos
    parser: argparse.ArgumentParser = argparse.ArgumentParser(
        prog="Sistema Val",
        description="""Sistema de valoração automática não assistida em lote.\n
        Selecione uma opção de valoração e o contrato a ser utilizado. Caso deseje, escolha uma família específica.\n
        Por padrão, o programa irá valoração todas as ordens pendentes das famílias:\n
        CAVALETE, HIDROMETRO, POCO, RAMAL AGUA, RELIGACAO, SUPRESSAO\n
        Para desobstrução não explicite contrato ou família.\n
        Não esqueça de colocar a senha para iniciar o programa.\n
        Opcional: habilite o Revalorar para reprocessar uma Ordem.\n
        Opções:\n
        1) Desvalorador\n
        2) Retrabalho\n
        3) N3 - Pertence ao serviço principal\n
        4) TSEs pré-selecionais como 'carteira de serviços'\n
        5) TSEs específicas digitadas pelo usuário\n
        6) TSE específica digitada pelo usuário\n
        7) Ordem específica digitada pelo usuário\n
        8) Pendentes em CSV ou Excel\n
        9) Por Família\n
        10) Desobstrução NORTESUL""",
        epilog="Author: Hyslan Silva Cruz <hyslansilva@gmail.com>",
    )
    parser.add_argument(
        "-s",
        "--session",  # It's a connection not Session
        type=int,
        default=0,
        help="Número da sessão do SAP a ser utilizada.",
        required=True,
    )
    parser.add_argument(
        "-o",
        "--option",
        type=str,
        help="Escolha uma Opção de valoração.",
        choices=[str(i) for i in range(1, 11)],
        required=True,
    )
    parser.add_argument("-c", "--contrato", default=None, type=str, help="Escolha o Contrato a ser utilizado.")
    parser.add_argument("-f", "--family", default=None, type=str, nargs="+", help="Escolha a Família a ser utilizada.")
    parser.add_argument("-p", "--password", type=str, help="Digite a senha para iniciar o programa.", required=True)
    parser.add_argument("-r", "--revalorar", default=False, type=bool, help="Revalorar uma Ordem.")

    args: argparse.Namespace = parser.parse_args()

    options: str = args.option
    # Avatar.
    val_avatar()
    # Inicialização
    hello(console)
    load_dotenv()

    if args.password != os.environ["PWD"]:
        console.print("Senha incorreta!\n Você não vai passar!. :mage:", style="bold")
        you_cant_pass("video")
        return

    # * Conexão ao SAP
    token, session = establish_sap_connection(args, console)
    # * Processamento das opções
    process_valoration(options, args, session, token, console)

    # Encerramento
    console.print("- Val: Valoração finalizada, encerrando. :star:")
    sap.fechar_conexao(args.session)
    logger.info("Valoração finalizada.")
    return


def establish_sap_connection(args: argparse.Namespace, console: rich.console.Console) -> tuple[str, win32com.client.CDispatch]:
    """Obtém o token de acesso e a sessão do SAP GUI.

    Args:
    ----
        args (argparse.Namespace): Número da Conexão.
        console (rich.console.Console): Terminal Highligthing

    Returns:
    -------
        tuple[str, win32com.client.CDispatch]: Token SSO e Sessão do SAP GUI.

    """
    try:
        console.print("[i cyan] Conectando ao SAP GUI e obtendo token de acesso...")
        token = down_sap()
        time.sleep(10)
        count_conexoes = sap.count_connections()
        session = None

        while count_conexoes < args.session:
            if count_conexoes == args.session:
                console.print(f"[bold green] Termo de conexão {args.session} atingido, obtendo conexão...")
                session = sap.choose_connection(args.session)
                break

            console.print(f"[bold yellow] Termo de conexão {args.session} não atingido, obtendo nova conexão...")
            token = sap.get_connection(token)
            time.sleep(10)
            count_conexoes = sap.count_connections()

        if session is None:
            session = sap.choose_connection(args.session)

        if session is None:
            msg = "Erro ao obter a sessão do SAP GUI."
            raise RuntimeError(msg)

        console.print("Conexão obtida com sucesso!")
    except pywintypes.com_error as con_error:
        msg = "Erro inesperado ao conectar ao SAP GUI."
        raise RuntimeError(msg) from con_error

    return token, session


def hello(console: rich.console.Console) -> None:
    """Imprime a saudação inicial do programa.

    Args:
    ----
        console (rich.console.Console): Terminal Highligthing

    """
    console.print("\n[bold blue underline]Sistema Val[/bold blue underline] :smiley:", justify="full")
    # Obtém a hora atual
    fuso_horario_sp = pytz.timezone("America/Sao_Paulo")
    hora_atual: datetime.time = fuso_horario_sp.localize(datetime.datetime.now()).time()
    hora: int = hora_atual.hour
    manha: int = 12
    tarde: int = 18
    saudacao = "Bom dia!" if hora < manha else "Boa tarde!" if hora < tarde else "Boa noite!"
    console.print(f"- Val: {saudacao} :star:\n- Val: Como vai você hoje?")
    console.print(f"- Val: Hora atual: {hora_atual.strftime('%H:%M:%S')} :alarm_clock:")


# -Main------------------------------------------------------------------
if __name__ == "__main__":
    main()
# -End-------------------------------------------------------------------
