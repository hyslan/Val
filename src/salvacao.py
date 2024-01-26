# salvacao.py
'''Módulo de salvar valoração.'''
import sys
import re
import threading
import pywintypes
from src import sql_view
from src.sap_connection import connect_to_sap
from src.confere_os import consulta_os
from src.sap import encerrar_sap


# Adicionando um Lock
lock = threading.Lock()


def salvar(ordem, qtd_ordem, contrato, unadm):
    '''Salvar e verificar se está salvando.'''

    def salvar_valoracao():
        '''Função para salvar valoração.'''
        nonlocal ordem
        nonlocal qtd_ordem
        nonlocal contrato
        nonlocal unadm

        # Seção Crítica - uso do Lock
        with lock:
            session = connect_to_sap()
            try:
                print("Salvando valoração!")
                session.findById("wnd[0]").sendVKey(11)
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
                rodape = session.findById("wnd[0]/sbar").Text  # Rodapé
                salvo = "Ajustes de valoração salvos com sucesso."
                if salvo == rodape:
                    print(f"{ordem} salva!")
                else:
                    rodape = rodape.lower()
                    padrao = r"material (\d+)"
                    correspondencias = re.search(padrao, rodape)
                    if correspondencias:
                        # Group 1 retira string 'material'
                        codigo_material = correspondencias.group(1)
                        print(codigo_material)
                        sys.exit()
            # pylint: disable=E1101
            except pywintypes.com_error:
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
                rodape = session.findById("wnd[0]/sbar").Text  # Rodapé
                rodape = rodape.lower()
                padrao = r"material (\d+)"
                correspondencias = re.search(padrao, rodape)
                if correspondencias:
                    # Group 1 retira string 'material'
                    codigo_material = correspondencias.group(1)
                    print(codigo_material)
                    sys.exit()

                print(f"Ordem: {ordem} não foi salva.")
                ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
                ja_valorado.valorada(obs="Não foi salvo")

        # Verificar se Salvou
        (status_sistema,
            status_usuario,
            *_) = consulta_os(ordem, contrato, unadm)
        print("Verificando se Ordem foi valorada.")
        if status_usuario == "EXEC VALO":
            print(f"Status da Ordem: {status_sistema}, {status_usuario}")
            print("Foi Salvo com sucesso!")
            ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
            ja_valorado.valorada("SIM")
            # Incremento + de Ordem.
            qtd_ordem += 1
        else:
            print(f"Ordem: {ordem} não foi salva.")
            ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
            ja_valorado.valorada(obs="Não foi salvo")

    # Start
    thread = threading.Thread(target=salvar_valoracao)
    thread.start()
    # Aguarde a thread concluir
    thread.join(timeout=300)
    if thread.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()

    return qtd_ordem
