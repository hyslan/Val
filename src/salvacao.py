# salvacao.py
'''Módulo de salvar valoração.'''
import sys
import re
import threading
import pythoncom
import win32com.client as win32
import pywintypes
from src import sql_view
from src.confere_os import consulta_os
from src.sap import Sap


# Adicionando um Lock
lock = threading.Lock()


def salvar(ordem, qtd_ordem, contrato, session):
    '''Salvar e verificar se está salvando.'''
    sap = Sap()

    def salvar_valoracao(session_id):
        '''Função para salvar valoração.'''
        nonlocal ordem
        nonlocal qtd_ordem
        nonlocal contrato
        nonlocal session

        # Seção Crítica - uso do Lock
        with lock:
            try:
                # pylint: disable=E1101
                pythoncom.CoInitialize()
                # pylint: disable=E1101
                gui = win32.Dispatch(
                    pythoncom.CoGetInterfaceAndReleaseStream(
                        session_id, pythoncom.IID_IDispatch)
                )
                print("Salvando valoração!")
                gui.findById("wnd[0]").sendVKey(11)
                gui.findById("wnd[1]/usr/btnBUTTON_1").press()
                rodape = gui.findById("wnd[0]/sbar").Text  # Rodapé
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
                gui.findById("wnd[1]/usr/btnBUTTON_1").press()
                rodape = gui.findById("wnd[0]/sbar").Text  # Rodapé
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
            *_) = consulta_os(ordem, session, contrato)
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

    # pylint: disable=E1101
    pythoncom.CoInitialize()
    session_id = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, session)
    # Start
    thread = threading.Thread(target=salvar_valoracao, kwargs={
                              'session_id': session_id})
    thread.start()
    # Aguarde a thread concluir
    thread.join(timeout=300)
    if thread.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        sap.encerrar_sap()

    return qtd_ordem
