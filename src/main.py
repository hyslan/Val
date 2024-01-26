'''Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz'''
# main.py
# Bibliotecas
import sys
import os
import time
import datetime
import getpass
from dotenv import load_dotenv
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from src import sql_view
from src.core import val
from src.avatar import val_avatar
from src.face_the_gandalf import you_cant_pass
from src.etl import extract_from_sql, pendentes_excel
from src.desvalorador import desvalorador
from src.retrabalhador import retrabalho
from src.osn3 import pertencedor


def contratada():
    '''input de contrato.'''
    contrato_gbitaquera = "4600042888"
    contrato_novasp = "4600041302"
    contrato_recape = "4600044782"
    contrato_nortesul = "4600043760"
    entrada = input("- Val: Qual o contrato?\n")
    if entrada == contrato_gbitaquera or entrada in ("GB", "GBITAQUERA", "gb"):
        contrato = contrato_gbitaquera
        unadm = "340"
    elif entrada == contrato_novasp or entrada in ("NOVASP", "novasp"):
        contrato = contrato_novasp
        unadm = "344"
    elif entrada == contrato_recape or entrada in ("RECAPE", "recape"):
        contrato = contrato_recape
        unadm = "344"
    elif entrada == contrato_nortesul or entrada in ("NORTESUL", "nortesul",
                                                     "NORTE SUL", "norte sul"):
        contrato = contrato_nortesul
        unadm = "344"
    else:
        print("Contrato não informado, encerrando.")
        sys.exit()

    return contrato, unadm


def main():
    '''Sistema principal da Val e inicializador do programa'''
    hora_parada = datetime.time(21, 50)  # Ponto de parada às 21h50min
    hora_retomada = datetime.time(6, 0)  # Ponto de retomada às 6h
    console = Console()
    table = Table(show_header=True, header_style="bold magenta",
                  title=":star: Menu de Opçôes :star:", style="bold")
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
            sys.exit()

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

        console.print(table)

        try:
            resposta = input(
                "\n- Val: Escolha uma opção:\n "
            )
            match resposta:
                case "1":
                    contrato, unadm = contratada()
                    desvalorador(contrato)
                    validador = True
                case "2":
                    contrato, unadm = contratada()
                    retrabalho(contrato, unadm)
                    validador = True
                case "3":
                    contrato, unadm = contratada()
                    pertencedor(contrato, unadm)
                    validador = True
                case "4":
                    contrato, unadm = contratada()
                    pendentes_list = extract_from_sql(contrato)
                    ordem, int_num_lordem, validador = val(
                        pendentes_list, contrato, unadm)
                case "5":
                    contrato, unadm = contratada()
                    tses_existentes = sql_view.Tabela("", "")
                    console.print("\n", tses_existentes.show_tses(),
                                  style="italic blue", justify="full")
                    tse_expec = input(
                        "- Val: Digite as TSE separadas por vírgula, por favor.\n")
                    lista_tse = tse_expec.split(', ')
                    pendentes = sql_view.Tabela(ordem="", cod_tse=lista_tse)
                    pendentes_list = pendentes.tse_escolhida(contrato)
                    ordem, int_num_lordem, validador = val(
                        pendentes_list, contrato, unadm)
                case "6":
                    contrato, unadm = contratada()
                    tses_existentes = sql_view.Tabela("", "")
                    console.print("\n", tses_existentes.show_tses(),
                                  style="italic blue", justify="full")
                    tse_expec = input(
                        "- Val: Digite a TSE expecífica, por favor.\n")
                    pendentes = sql_view.Tabela(ordem="", cod_tse=tse_expec)
                    pendentes_list = pendentes.tse_expecifica(contrato)
                    ordem, int_num_lordem, validador = val(
                        pendentes_list, contrato, unadm)
                case "7":
                    contrato, unadm = contratada()
                    ordem_expec = input(
                        "- Val: Digite o Nº da Ordem, por favor.\n"
                    )
                    teste = sql_view.Tabela(ordem_expec, "")
                    teste_list = teste.ordem_especifica(contrato)
                    ordem, int_num_lordem, validador = val(
                        teste_list, contrato, unadm
                    )
                case "8":
                    contrato, unadm = contratada()
                    planilha = pendentes_excel()
                    ordem, int_num_lordem, validador = val(
                        planilha, contrato, unadm)
                case "9":
                    contrato, unadm = contratada()
                    pendentes = sql_view.Tabela("", "_")
                    console.print(Panel.fit(
                        pendentes.show_family()), style="italic yellow" "\n")
                    familia = input(
                        "- Val: Digite o nome da família, por favor.\n"
                    )
                    pendentes_list = pendentes.familia(familia, contrato)
                    ordem, int_num_lordem, validador = val(
                        pendentes_list, contrato, unadm
                    )

        except TypeError as erro:
            print(f"Erro: {erro}")
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
                        f"- Val: Retomando Ordem: {ordem} \n da linha: {int_num_lordem}")
                    break  # Sai do Loop quando atingir a hora de retomada.


# -Main------------------------------------------------------------------
if __name__ == '__main__':
    main()
# -End-------------------------------------------------------------------
