"""Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz"""
# main.py
# Bibliotecas
import sys
import os
import time
import datetime
import getpass
import pywintypes
import numpy as np
from dotenv import load_dotenv
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
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


def contratada():
    """Input do contrato"""
    contratos_mlg = {

        "NOVASP": {"contrato": "4600041302", "unadm": "344", "municipio": "100"},
        "RECAPE": {"contrato": "4600044782", "unadm": "344", "municipio": "100"},
        "NORTESUL": {"contrato": "4600043760", "unadm": "344", "municipio": "100"},
    }
    contratos_mlq = {
        "GBITAQUERA": {"contrato": "4600042888", "unadm": "340", "municipio": "100"},
        "NOVASP": {"contrato": "4600041302", "unadm": "340", "municipio": "100"},
        "NORTESUL": {"contrato": "4600043654", "unadm": "340", "municipio": "100"},
        "RECAPE": {"contrato": "4600044777", "unadm": "340", "municipio": "100"},
        "ZIGURATE": {"contrato": "4600042975", "unadm": "340", "municipio": "100"},
    }
    contratos_mln = {
        "ZIGURATE": {"contrato": "4600042975", "unadm": "348", "municipio": ""},
        "NORTESUL": {"contrato": "4600045267", "unadm": "348", "municipio": ""},
        "NORTESUL2": {"contrato": "4600046036", "unadm": "348", "municipio": ""},
        "RECAPE": {"contrato": "4600044787", "unadm": "348", "municipio": ""}
    }
    ugrs = {
        "MLG": contratos_mlg,
        "MLQ": contratos_mlq,
        "MLN": contratos_mln
    }

    regiao = input("- Val: Qual UGR:\n").upper()
    if regiao in ugrs:
        empresa = input("- Val: Qual o contrato?\n").upper()
        if empresa in ugrs[regiao]:
            contrato_info = ugrs[regiao][empresa]
            contrato, unadm, municipio = contrato_info[
                "contrato"], contrato_info["unadm"], contrato_info["municipio"]
        else:
            print("Contrato não informado, encerrando.")
            sys.exit(0)
    else:
        print("Não encontrei a UGR")
        sys.exit(0)

    return contrato, municipio


def main():
    """Sistema principal da Val e inicializador do programa"""
    validador = False
    ordem = ""
    hora_parada = datetime.time(21, 50)  # Ponto de parada às 21h50min
    hora_retomada = datetime.time(6, 0)  # Ponto de retomada às 6h
    console = Console()
    table = Table(show_header=True, header_style="bold magenta",
                  title=":star: Menu de Opçôes :star:", style="bold")
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
        init = getpass.getpass("Digite a senha por favor.\n")
        if not init == os.environ["pwd"]:
            console.print(
                "Senha incorreta!\n Você não vai passar!. :mage:", style="bold")
            you_cant_pass()
            sys.exit(0)

        table.add_column("Seletor", style="dim", width=12)
        table.add_column("Tipo")
        table.add_row("1", "Desvalorador")
        table.add_row("2", "Retrabalho")
        table.add_row("3", "Pertencedor")
        table.add_row("4", "TSE geral")
        table.add_row("5", "TSEs Expecíficas")
        table.add_row("6", "TSE Expecífica")
        table.add_row("7", "Teste de Ordem única")
        table.add_row("8", "Planilha de pendentes")
        table.add_row("9", "Família de serviço")

        try:
            session = sap.escolher_sessao()
        # pylint: disable=E1101
        except pywintypes.com_error:
            console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
            console.print(
                "[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
            down_sap()
            session = sap.escolher_sessao()

        console.print(table)

        try:
            resposta = input(
                "\n- Val: Escolha uma opção:\n "
            )
            match resposta:
                case "1":
                    contrato = contratada()
                    desvalorador(contrato, session)
                    validador = True
                case "2":
                    contrato = contratada()
                    retrabalho(contrato, session)
                    validador = True
                case "3":
                    contrato = contratada()
                    pertencedor(contrato, session)
                    validador = True
                case "4":
                    contrato = contratada()
                    pendentes_list = extract_from_sql(contrato[0])
                    ordem, validador = val(
                        pendentes_list, session, contrato)
                case "5":
                    contrato = contratada()
                    tses_existentes = sql_view.Tabela("", "")
                    console.print("\n", tses_existentes.show_tses(),
                                  style="italic blue", justify="full")
                    tse_expec = input(
                        "- Val: Digite as TSE separadas por vírgula, por favor.\n")
                    lista_tse = tse_expec.split(', ')
                    pendentes = sql_view.Tabela(ordem="", cod_tse=lista_tse)
                    pendentes_array = pendentes.tse_escolhida(contrato[0])
                    ordem, validador = val(
                        pendentes_array, session, contrato)
                case "6":
                    contrato = contratada()
                    tses_existentes = sql_view.Tabela("", "")
                    console.print("\n", tses_existentes.show_tses(),
                                  style="italic blue", justify="full")
                    tse_expec = input(
                        "- Val: Digite a TSE expecífica, por favor.\n")
                    pendentes = sql_view.Tabela(ordem="", cod_tse=tse_expec)
                    pendentes_array = pendentes.tse_expecifica(contrato[0])
                    ordem, validador = val(
                        pendentes_array, session, contrato)
                case "7":
                    contrato = contratada()
                    ordem_expec = input(
                        "- Val: Digite o Nº da Ordem, por favor.\n"
                    )
                    mun = input("Digite o Nº do Município.\n")
                    pendentes_array = np.array([[ordem_expec, mun]])
                    ordem, validador = val(
                        pendentes_array, session, contrato
                    )
                case "8":
                    contrato = contratada()
                    ask = input("é csv?")
                    if ask == "s":
                        planilha = pendentes_csv
                    else:
                        planilha = pendentes_excel()

                    ordem, validador = val(
                        planilha, session, contrato)
                case "9":
                    contrato = contratada()
                    pendentes = sql_view.Tabela("", "_")
                    console.print(Panel.fit(
                        pendentes.show_family()), style="italic yellow" "\n")
                    familia = input(
                        "- Val: Digite o nome da família, por favor.\n"
                    )
                    pendentes_array = pendentes.familia(familia, contrato[0])
                    ordem, validador = val(
                        pendentes_array, session, contrato
                    )

        except (TypeError, ValueError) as erro:
            console.print(f":fearful: Erro Main: {erro}", style="bold red")
            console.print_exception()
            return

        if validador is True:
            print("- Val: Valoração finalizada, encerrando.")
            sys.exit()

        # Loop de Parada
        if hora_atual >= hora_parada:
            print("A Val foi descansar.")
            print("- Val: até amanhã.")
            while True:
                hora_atual = datetime.datetime.now().time()
                print(f"Hora atual na parada: {hora_atual}")
                time.sleep(60)
                if hora_atual >= hora_retomada:
                    print("- Val: Bom dia!")
                    print("- Val: Vamos trabalhar!")
                    print(
                        f"- Val: Retomando Ordem: {ordem} \n")
                    break  # Sai do Loop quando atingir a hora de retomada.


# -Main------------------------------------------------------------------
if __name__ == '__main__':
    main()
# -End-------------------------------------------------------------------
