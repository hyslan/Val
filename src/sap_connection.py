# sap_connection.py
# -Begin-----------------------------------------------------------------
"""Módulo SAP"""

# -Bibliotecas--------------------------------------------------------------
import win32com.client
import pythoncom
import pywintypes
import time
from rich.console import Console
from src.sapador import down_sap

# -Sub Main--------------------------------------------------------------
console = Console()


def populate_sessions() -> None:
    def connect_to_sap():
        """Função para conexão SAP"""
        pythoncom.CoInitialize()
        sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        conn = application.Children(0)
        sess = conn.Children

        return sess, conn

try:
    session, connection = connect_to_sap()
# pylint: disable=E1101
except pywintypes.com_error:
    console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
    console.print(
        "[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
    down_sap()
    time.sleep(20)
    session = connect_to_sap()

    # Obtendo o índice da última sessão ativa
    ultimo_indice = len(session) - 1
    if ultimo_indice < 5:
        for _ in range(ultimo_indice, 5):
            connection.Children(ultimo_indice).CreateSession()
