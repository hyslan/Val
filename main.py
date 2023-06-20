'''Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz'''
# main.py
# Bibliotecas
import sys
import time
import datetime
from core import val
from avatar import val_avatar


def main():
    '''Sistema principal da Val e inicializador do programa'''
    caminho_avatar = 'C:/Users/irgpapais/Documents/Meus Projetos/val/val.png'
    hora_parada = datetime.time(21, 50)  # Ponto de parada às 21h50min
    hora_retomada = datetime.time(6, 0)  # Ponto de retomada às 6h
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
        # Início do Sistema
        ordem, int_num_lordem, validador = val()
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
