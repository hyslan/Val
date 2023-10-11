'''Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz'''
# main.py
# Bibliotecas
import sys
import time
import datetime
import sql_view
from core import val
from avatar import val_avatar
from etl import extract_from_sql
from desvalorador import desvalorador
from retrabalhador import retrabalho


def main():
    '''Sistema principal da Val e inicializador do programa'''
    caminho_avatar = 'C:/Users/irgpapais/Documents/Meus Projetos/val/val.png'
    hora_parada = datetime.time(21, 50)  # Ponto de parada às 21h50min
    hora_retomada = datetime.time(6, 0)  # Ponto de retomada às 6h
    # Avatar.
    val_avatar(caminho_avatar)

    while True:
        print("\nSistema Val.")
        # Obtém a hora atual
        hora_atual = datetime.datetime.now().time()
        hora = hora_atual.hour
        if hora < 12:
            saudacao = "Bom dia!"
        elif hora < 18:
            saudacao = "Boa tarde!"
        else:
            saudacao = "Boa noite!"
        print(f"- Val: {saudacao}\n- Val: Como vai você hoje?")
        print(f"\n- Val: Hora atual: {hora_atual}")

        print("\nMenu de Opções")
        print("\n1 - Desvalorador"
              + "\n2 - Retrabalho"
              + "\n3 - TSE geral"
              + "\n4 - TSEs Expecíficas"
              + "\n5 - TSE Expecífica")
        try:
            resposta = input(
                "\n- Val: Escolha uma opção:\n "
            )
            match resposta:
                case "1":
                    desvalorador()
                    validador = True
                case "2":
                    retrabalho()
                    validador = True
                case "3":
                    pendentes_list = extract_from_sql()
                    ordem, int_num_lordem, validador = val(pendentes_list)
                case "4":
                    tse_expec = input(
                        "- Val: Digite as TSE separadas por vírgula, por favor.\n")
                    lista_tse = tse_expec.split(', ')
                    pendentes = sql_view.Tabela(ordem="", cod_tse=lista_tse)
                    pendentes_list = pendentes.tse_escolhida()
                    ordem, int_num_lordem, validador = val(pendentes_list)
                case "5":
                    tse_expec = input(
                        "- Val: Digite a TSE expecífica, por favor.\n")
                    pendentes = sql_view.Tabela(ordem="", cod_tse=tse_expec)
                    pendentes_list = pendentes.tse_expecifica()
                    ordem, int_num_lordem, validador = val(pendentes_list)

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
