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
from src.temporizador import cronometro_val


# Adicionando um Lock
lock = threading.Lock()
console = Console()


def salvar(ordem, qtd_ordem, contrato, session, principal_tse, cod_mun):
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

                print(f"Ordem: {ordem} não foi salva.")
                time_spent = cronometro_val(start_time, ordem)
                ja_valorado = sql_view.Tabela(
                    ordem=ordem, cod_tse=principal_tse)
                ja_valorado.valorada(
                    obs=rodape,
                    valorado="NÃO", contrato=contrato, municipio=cod_mun,
                    status="DISPONÍVEL", data_valoracao=None,
                    matricula='117615', valor_medido=0, tempo_gasto=time_spent)

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
        time_spent = cronometro_val(start_time, ordem)
        ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse=principal_tse)
        ja_valorado.valorada(
            obs=rodape,
            valorado="SIM", contrato=contrato, municipio=cod_mun,
            status="VALORADA", data_valoracao=None,
            # TODO: get total amount from SAP
            matricula='117615', valor_medido=0, tempo_gasto=time_spent
        )
        # Incremento + de Ordem.
        qtd_ordem += 1
    else:
        console.print(
            f"Ordem: {ordem} não foi salva. :pouting_face:", style="italic yellow")
        time_spent = cronometro_val(start_time, ordem)
        ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse=principal_tse)
        ja_valorado.valorada(
            obs=rodape,
            valorado="NÃO", contrato=contrato, municipio=cod_mun,
            status="DISPONÍVEL", data_valoracao=None,
            matricula='117615', valor_medido=0, tempo_gasto=time_spent
        )

    ja_valorado.clean_duplicates()

    return qtd_ordem, rodape
