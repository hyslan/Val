# salvacao.py
"""Módulo de salvar valoração."""
import re
import threading
import pythoncom
import win32com.client as win32
import pywintypes
from rich.console import Console
from python.src import sql_view
from python.src import sap
from python.src.temporizador import cronometro_val


# Adicionando um Lock
lock = threading.Lock()
console = Console()


def salvar(ordem, qtd_ordem, contrato, session,
           principal_tse, cod_mun, start_time, n_con):
    """Salvar e verificar se está salvando."""
    salvo = "Ajustes de valoração salvos com sucesso."
    total = session.findById(
        "wnd[0]/usr/txtGS_HEADER-VAL_ATUAL").Text
    if not total == '':
        f_total = float(total.replace(".", "").replace(",", "."))
    else:
        f_total = 0

    def salvar_valoracao(session_id):
        """Função para salvar valoração."""
        nonlocal ordem, salvo
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
                # Rodapé

            # pylint: disable=E1101
            except pywintypes.com_error:
                gui.findById("wnd[0]").sendVKey(11)
                gui.findById("wnd[1]/usr/btnBUTTON_1").press()

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
            sap.fechar_conexao(n_con)

        # Check the footer.
        rodape = session.findById("wnd[0]/sbar").Text
        if salvo == rodape:
            console.print(f"[italic green] Ordem: {
                          ordem} Foi Salvo com sucesso! :rocket:")
            time_spent = cronometro_val(start_time, ordem)
            ja_valorado = sql_view.Sql(ordem=ordem, cod_tse=principal_tse)
            ja_valorado.valorada(
                obs=rodape,
                valorado="SIM", contrato=contrato, municipio=cod_mun,
                status="VALORADA", data_valoracao=None,
                matricula='117615', valor_medido=f_total, tempo_gasto=time_spent
            )
            # Incremento + de Ordem.
            qtd_ordem += 1
        else:
            console.print(f"[italic red]Nota de rodapé: {rodape}")
            rodape = rodape.lower()
            padrao = r"material (\d+)"
            correspondencias = re.search(padrao, rodape)
            if correspondencias:
                # Group 1 retira string 'material'
                codigo_material = correspondencias.group(1)
                print(codigo_material)

            console.print(
                f"Ordem: {ordem} não foi salva.", style="italic red")
            console.print(f"[bold yellow]Motivo: {rodape}")
            time_spent = cronometro_val(start_time, ordem)
            ja_valorado = sql_view.Sql(
                ordem=ordem, cod_tse=principal_tse)
            ja_valorado.valorada(
                obs=rodape,
                valorado="NÃO", contrato=contrato, municipio=cod_mun,
                status="DISPONÍVEL", data_valoracao=None,
                matricula='117615', valor_medido=0, tempo_gasto=time_spent)

        ja_valorado.clean_duplicates()
    except pywintypes.com_error as salvar_erro:
        console.print(f"Erro na parte de salvar: {salvar_erro} :pouting_face:",
                      style="bold red")

    return qtd_ordem, rodape