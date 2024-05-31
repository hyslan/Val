#!/usr/bin/env python3
"""Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz"""
# main.py
# Bibliotecas
import argparse
import os
import time
import datetime
import getpass
import pywintypes
import numpy as np
from dotenv import load_dotenv
from rich.console import Console
from src.sap import Sap
from src import sql_view
from src.core import val
from src.avatar import val_avatar
from src.face_the_gandalf import you_cant_pass
from src.etl import extract_from_sql, pendentes_excel, pendentes_csv
from src.desvalorador import desvalorador
from src.retrabalhador import retrabalho
from src.osn3 import pertencedor
from src.sapador import down_sap


def main():
    """Sistema principal da Val e inicializador do programa"""
    parser = argparse.ArgumentParser(prog="Sistema Val",
                                     description="Sistema de valoração automática não assistida.",
                                     epilog="Author: Hyslan Silva Cruz")
    parser.add_argument('-s', '--session',
                        type=int, default=0,
                        help='Número da sessão do SAP a ser utilizada.')
    parser.add_argument('-o', '--option',
                        type=str, help="Escolha uma Opção de valoração.",
                        choices=[str(i) for i in range(1, 10)])
    parser.add_argument('-c', '--contrato',
                        type=str, help="Escolha o Contrato a ser utilizado.")
    parser.add_argument('-f', '--family', default='hidrometro',
                        type=str, help="Escolha a Família a ser utilizada.")
    parser.add_argument('-p', '--password',
                        type=str, help="Digite a senha para iniciar o programa.")

    args = parser.parse_args()
    options = args.option
    validador = False
    hora_parada = datetime.time(21, 50)  # Ponto de parada às 21h50min
    console = Console()
    sap = Sap()
    # Avatar.
    val_avatar()

    while True:
        console.print(
            "\n[bold blue underline]Sistema Val[/bold blue underline] :smiley:", justify='full')
        # Obtém a hora atual
        hora_atual = datetime.datetime.now().time()
        hora = hora_atual.hour
        if hora < 12:
            saudacao = "Bom dia!"
        elif hora < 18:
            saudacao = "Boa tarde!"
        else:
            saudacao = "Boa noite!"
        console.print(f"- Val: {saudacao} :star:\n- Val: Como vai você hoje?")
        console.print(
            f"- Val: Hora atual: {hora_atual.strftime('%H:%M:%S')} :alarm_clock:")
        load_dotenv()

        if not args.password == os.environ["pwd"]:
            console.print(
                "Senha incorreta!\n Você não vai passar!. :mage:", style="bold")
            you_cant_pass('video')
            return

        try:
            session = sap.escolher_sessao(args.session)
        # pylint: disable=E1101
        except pywintypes.com_error:
            console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
            console.print(
                "[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
            down_sap()
            session = sap.escolher_sessao(args.session)

        try:
            match options:
                case "1":
                    desvalorador(args.contrato, session)
                    validador = True
                case "2":
                    retrabalho(args.contrato, session)
                    validador = True
                case "3":
                    pertencedor(args.contrato, session)
                    validador = True
                case "4":
                    pendentes_list = extract_from_sql(args.contrato)
                    ordem, validador = val(
                        pendentes_list, session, args.contrato)
                case "5":
                    tses_existentes = sql_view.Tabela("", "")
                    console.print("\n", tses_existentes.show_tses(),
                                  style="italic blue", justify="full")
                    tse_expec = input(
                        "- Val: Digite as TSE separadas por vírgula, por favor.\n")
                    lista_tse = tse_expec.split(', ')
                    pendentes = sql_view.Tabela(ordem="", cod_tse=lista_tse)
                    pendentes_array = pendentes.tse_escolhida(args.contrato)
                    ordem, validador = val(
                        pendentes_array, session, args.contrato)
                case "6":
                    tses_existentes = sql_view.Tabela("", "")
                    console.print("\n", tses_existentes.show_tses(),
                                  style="italic blue", justify="full")
                    tse_expec = input(
                        "- Val: Digite a TSE expecífica, por favor.\n")
                    pendentes = sql_view.Tabela(ordem="", cod_tse=tse_expec)
                    pendentes_array = pendentes.tse_expecifica(args.contrato)
                    ordem, validador = val(
                        pendentes_array, session, args.contrato)
                case "7":
                    ordem_expec = input(
                        "- Val: Digite o Nº da Ordem, por favor.\n"
                    )
                    mun = input("Digite o Nº do Município.\n")
                    pendentes_array = np.array([[ordem_expec, mun]])
                    ordem, validador = val(
                        pendentes_array, session, args.contrato
                    )
                case "8":
                    ask = input("é csv?")
                    if ask == "s":
                        planilha = pendentes_csv
                    else:
                        planilha = pendentes_excel()

                    ordem, validador = val(
                        planilha, session, args.contrato)
                case "9":
                    pendentes = sql_view.Tabela("", "_")
                    pendentes_array = pendentes.familia(args.family, args.contrato)
                    ordem, validador = val(
                        pendentes_array, session, args.contrato
                    )

        except (TypeError, ValueError) as erro:
            console.print(f":fearful: Erro Main: {erro}", style="bold red")
            console.print_exception()
            return

        if validador is True:
            print("- Val: Valoração finalizada, encerrando.")
            return

        # Loop de Parada
        if hora_atual >= hora_parada:
            print("A Val foi descansar.")
            print("- Val: até amanhã.")
            return


# -Main------------------------------------------------------------------
if __name__ == '__main__':
    main()
# -End-------------------------------------------------------------------
