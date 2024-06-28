# salvacao.py
"""Módulo de salvar valoração."""
import sys
import re
import threading
import pythoncom
import win32com.client as win32
import pywintypes
from rich.console import Console
from src import sql_view
from src.confere_os import consulta_os
from src import sap


# Adicionando um Lock
lock = threading.Lock()
console = Console()


def salvar(ordem, qtd_ordem, contrato, session):
    """Salvar e verificar se está salvando."""
    rodape = None
    salvo = "Ajustes de valoração salvos com sucesso."

    def salvar_valoracao(session_id):
        """Função para salvar valoração."""
        nonlocal ordem, rodape, salvo

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

                if salvo == rodape:
                    print(f"{ordem} salva!")
                else:
                    console.print(f"[italic red]Nota de rodapé: {rodape}")
                    rodape = rodape.lower()
                    padrao = r"material (\d+)"
                    correspondencias = re.search(padrao, rodape)
                    if correspondencias:
                        # Group 1 retira string 'material'
                        codigo_material = correspondencias.group(1)
                        print(codigo_material)

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

            return rodape
    try:
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

        if not salvo == rodape:
            return qtd_ordem, rodape
    except pywintypes.com_error as salvar_erro:
        console.print(f"Erro na parte de salvar: {salvar_erro} :pouting_face:",
                      style="bold red")

    # Verificar se Salvou
    (status_sistema,
        status_usuario,
        *_) = consulta_os(ordem, session, contrato)
    print("Verificando se Ordem foi valorada.")
    if status_usuario == "EXEC VALO":
        print(f"Status da Ordem: {status_sistema}, {status_usuario}")
        console.print("[italic green]Foi Salvo com sucesso! :rocket:")
        ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
        ja_valorado.valorada("SIM")
        # Incremento + de Ordem.
        qtd_ordem += 1
    else:
        console.print(f"Ordem: {ordem} não foi salva. :pouting_face:", style="italic yellow")
        ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
        ja_valorado.valorada(obs="Não foi salvo")

    ja_valorado.clean_duplicates()

    return qtd_ordem, rodape
