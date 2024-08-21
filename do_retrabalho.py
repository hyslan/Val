"""Every second friday on the month, do this rpa for 'Retrabalho Confirmado' orders."""
import pywintypes
import rich.console
import win32com.client
from rich.console import Console

from python.src import sap
from python.src.retrabalhador import retrabalho
from python.src.sapador import down_sap

console: rich.console.Console = Console()


def do() -> None:
    try:
        # Always use the first session
        session: win32com.client.CDispatch = sap.choose_connection(0)
    # pylint: disable=E1101
    except pywintypes.com_error:
        console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
        console.print(
            "[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
        down_sap()
        session: win32com.client.CDispatch = sap.choose_connection(0)

    retrabalho("", session)


if __name__ == "__main__":
    do()
