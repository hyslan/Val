# sap_connection.py
# -Begin-----------------------------------------------------------------
"""Módulo SAP"""

# -Bibliotecas--------------------------------------------------------------
import pywintypes
import time
from rich.console import Console
from python.src.sapador import down_sap
import python.src.sap as sap

# -Sub Main--------------------------------------------------------------
console = Console()


def populate_sessions() -> None:
    def rise_connection(token: str) -> None:
        for _ in range(3):
            sap.get_connection(token)
            time.sleep(5)
    try:
        token = down_sap()
        print("Wait 20 seconds...")
        time.sleep(20)
        rise_connection(token)

        time.sleep(10)
    except pywintypes.com_error:
        console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
        console.print(
            "[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
        token = down_sap()
        print("Wait 20 seconds...")
        time.sleep(20)
        rise_connection(token)
# --- END
