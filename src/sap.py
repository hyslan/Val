"""Módulo para interagir com o SAP GUI"""
import subprocess
import win32com.client
import pythoncom
from rich.console import Console
import rich.console

console: rich.console.Console = Console()


def listar_conexoes() -> win32com.client.CDispatch:
    """Função para listar as conexões ativas."""
    # pylint: disable=E1101
    pythoncom.CoInitialize()
    sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
    application: win32com.client.CDispatch = sapguiauto.GetScriptingEngine
    connections: win32com.client.CDispatch = application.Children(0)

    return connections


def listar_sessoes() -> win32com.client.CDispatch:
    """Função para listar as sessions ativas"""
    # pylint: disable=E1101
    pythoncom.CoInitialize()
    sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
    application: win32com.client.CDispatch = sapguiauto.GetScriptingEngine
    connection: win32com.client.CDispatch = application.Children(0)
    sessions: win32com.client.CDispatch = connection.Children

    return sessions


def contar_sessoes() -> int:
    """Contar por tamanho de 1 a 6, caso for criar sessão subtrair -1"""
    # pylint: disable=E1101
    pythoncom.CoInitialize()
    sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
    application: win32com.client.CDispatch = sapguiauto.GetScriptingEngine
    connection: win32com.client.CDispatch = application.Children(0)
    sessions: win32com.client.CDispatch = connection.Children
    console.print(
        f"[blue italic]Quantidade de sessões ativas: {sessions.Count}")
    return sessions.Count


def criar_sessao(sessions) -> win32com.client.CDispatch:
    """Função para criar sessões"""
    # pylint: disable=E1101
    pythoncom.CoInitialize()
    sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
    application: win32com.client.CDispatch = sapguiauto.GetScriptingEngine
    connection: win32com.client.CDispatch = application.Children(0)

    # Obtendo o índice da última sessão ativa
    ultimo_indice = len(sessions) - 1

    # Criando uma nova sessão com base na última sessão ativa
    if ultimo_indice < 5:
        connection.Children(ultimo_indice).CreateSession()
        while ultimo_indice >= len(sessions) - 1:
            sessions = connection.Children

        # Acessando a nova sessão
        session = connection.Children(len(sessions) - 1)
    else:
        session = connection.Children(5)

    return session


def escolher_sessao(n_selected: int) -> win32com.client.CDispatch:
    """Escolher com qual sessão trabalhar"""
    # pylint: disable=E1101
    pythoncom.CoInitialize()
    sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
    application: win32com.client.CDispatch = sapguiauto.GetScriptingEngine
    connection: win32com.client.CDispatch = application.Children(0)
    session: win32com.client.CDispatch = connection.Children(n_selected)
    return session


def fechar_conexao() -> None:
    """Função para fechar o SAP."""
    # pylint: disable=E1101
    pythoncom.CoInitialize()
    sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
    application: win32com.client.CDispatch = sapguiauto.GetScriptingEngine
    connection: win32com.client.CDispatch = application.Children(0)
    connection.CloseConnection()


def encerrar_sap() -> None:
    """Encerra o app SAP"""
    # ? Can stop the SAP GUI without destroy the process?
    processo: str = 'saplogon.exe'
    try:
        subprocess.run(['taskkill', '/F', '/IM', processo], check=True)
        print(f'O processo {processo} foi encerrado com sucesso.')
    except subprocess.CalledProcessError:
        print(f'Não foi possível encerrar o processo {processo}.')


def get_connection():
    sap_access = (
    '[System]\n'
    'Name=EP0\n'
    'Client=100\n'
    r'GuiParm=/M/erpprdci.ti.sabesp.com.br/S/3908/G/PRODUCAO /UPDOWNLOAD_CP=1160'
    '\n'
    '[User]\n'
    'Name=117615\n'
    fr'at="MYSAPSSO2={token}"'
    '\n'
    'Language=PT\n'
    '[Function]\n'
    'Command=SMEN\n'
    'Type=Transaction\n'
    '[Configuration]\n'
    'Workplace=false\n'
    'GuiSize=\n'
    '[Options]\n'
    'Reuse=-1')
    with open('tx.sap', 'w') as s:
        s.write(sap_access)
    subprocess.run(['C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', '-system=EP0', '- client=100', '-user=117615', '-pw=123456', '-language=PT', '-command=SMEN', '-guiparm=/M/erpprdci.ti.sabesp.com.br/S/3908/G/PRODUCAO /UPDOWNLOAD_CP=1160', '-title=EP0 - 100 - 117615 - PT - SMEN', '-reuse=-1', '-nosplash', '-noconfig', '-noautoal', '-noupdinfo', '-nocrashdump', '-nologo', '-nofork', '-noini