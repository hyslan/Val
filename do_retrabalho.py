"""Every second friday on the month, do this rpa for 'Retrabalho Confirmado' orders."""
import win32com.client
import pywintypes
import rich.console
from rich.console import Console
import src.sap as sap
from src.sapador import down_sap
from src.retrabalhador import retrabalho

console: rich.console.Console = Console()


def do() -> None:
    try:
        # Always use the first session
        session: win32com.client.CDispatch = sap.escolher_sessao(0)
    # pylint: disable=E1101
    except pywintypes.com_error:
        console.print("[bold cyan] Ops! o SAP Gui não está aberto.")
        console.print(
            "[bold cyan] Executando o SAP GUI\n Por favor aguarde...")
        down_sap()
        session: win32com.client.CDispatch = sap.escolher_sessao(0)

    retrabalho("",session)


if __name__ == '__main__':
    do()
