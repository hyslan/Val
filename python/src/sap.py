"""Módulo para interagir com o SAP GUI."""

import contextlib
import logging
import os
import subprocess
import time
from pathlib import Path

import pythoncom
import pywintypes
import rich.console
import win32com.client
from dotenv import load_dotenv
from rich.console import Console

console: rich.console.Console = Console()
logger = logging.getLogger(__name__)


def reconnect(session: pywintypes.IID) -> pywintypes.HANDLE:  # type: ignore
    """Reconectar a sessão ao SAP GUI.

    Args:
    ----
        session (_type_): _description_

    Returns:
    -------
        pywintypes.HANDLE: _description_

    """
    console.print("Tentando obter de volta IID_IDispatch")
    s_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, session)
    console.print(f"Obtido com sucesso: {s_id}")
    return s_id


def connection_object(n_selected: int) -> win32com.client.CDispatch:
    """Função para listar as conexões ativas."""
    con = get_app()
    connection: win32com.client.CDispatch = con.Item(n_selected)
    return connection


def listar_sessoes(n_selected: int) -> win32com.client.CDispatch:
    """Função para listar as sessions ativas."""
    con = get_app()
    con_selected: win32com.client.CDispatch = con.Item(n_selected)
    sessions: win32com.client.CDispatch = con_selected.Sessions
    return sessions


def contar_sessoes(n_selected: int) -> int:
    """Contar por tamanho de 1 a 6, caso for criar sessão subtrair, então: -1."""
    con = get_app()
    con_selected: win32com.client.CDispatch = con.Item(n_selected)
    sessions: win32com.client.CDispatch = con_selected.Sessions
    console.print(f"[blue italic]Quantidade de sessões ativas: {sessions.Count}")
    return sessions.Count  # type: ignore


def create_session(n_selected: int) -> win32com.client.CDispatch:
    """Função para criar sessões."""
    con = get_app()
    con_selected: win32com.client.CDispatch = con.Item(n_selected)
    sessions = con_selected.Sessions
    # Obtendo o índice da última sessão ativa
    ultimo_indice = len(sessions) - 1

    # Criando uma nova sessão com base na última sessão ativa
    max_sessions = 5
    if ultimo_indice < max_sessions:
        con_selected.Children(ultimo_indice).CreateSession()
        while ultimo_indice >= len(sessions) - 1:
            sessions = con_selected.Children

        # Acessando a nova sessão
        session = con_selected.Children(len(sessions) - 1)
    else:
        session = con_selected.Children(5)

    return session  # type: ignore


def choose_connection(n_selected: int) -> win32com.client.CDispatch:
    """Escolher com qual sessão trabalhar."""
    con = get_app()
    session: win32com.client.CDispatch = con.Item(n_selected).Sessions(0)
    return session


def count_connections() -> int:
    """Contar conexões ativas."""
    con = get_app()
    return con.Count - 1  # type: ignore


def fechar_conexao(n_con: int) -> None:
    """Função para fechar o SAP."""
    con = get_app()
    connection = con.Item(n_con)
    connection.CloseConnection()


def encerrar_sap() -> None:
    """Encerra o app SAP."""
    # ? Can stop the SAP GUI without destroy the process?
    processo: str = "saplogon.exe"
    with contextlib.suppress(subprocess.CalledProcessError):
        subprocess.run(["taskkill", "/F", "/IM", processo], check=True)


def get_connection(token: str) -> str:
    """Cria nova conexão com o SAPGUI.

    Args:
    ----
        token (str): SSO token.

    Returns:
    -------
        str: SSO token.

    """
    load_dotenv()
    sap_access = (
        '[System]\n'
        'Name=EP0\n'
        'Client=100\n'
        rf'GuiParm={os.environ['SERVER']}'
        '\n'
        '[User]\n'
        f'Name={os.environ['USR']}\n'
        rf'at="MYSAPSSO2={token}"'
        '\n'
        'Language=PT\n'
        '[Function]\n'
        'Command=SMEN\n'
        'Type=Transaction\n'
        '[Configuration]\n'
        'Workplace=false\n'
        'GuiSize=\n'
        '[Options]\n'
        'Reuse=-1'
    )
    path_archive = str(Path.cwd() / "shortcut/repeat/tx.sap")
    with Path(path_archive).open("w") as s:
        s.write(sap_access)

    # Execute the command
    try:
        subprocess.run(["powershell", "start", f'"{path_archive}"'], check=False)
        time.sleep(5)
        # Verifica se o processo está em execução
        if not is_process_running("powershell.exe"):
            pass
    except subprocess.CalledProcessError:
        pass

    return token


def is_process_running(process_name: str) -> bool | None:
    """Verifica se o processo está em execução."""
    try:
        subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {process_name}"], capture_output=True, check=False)
    except subprocess.CalledProcessError:
        return False
    else:
        return True


def get_app() -> win32com.client.CDispatch:
    """Get the SAP GUI application.

    Returns
    -------
        win32com.client.CDispatch: Connections from application

    """
    pythoncom.CoInitialize()
    app: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    return app.Connections  # type: ignore


def keep_session_active(n_selected: int, token: str) -> None:
    """Keep the session active."""
    session = create_session(n_selected)
    while True:
        try:
            session.findById("wnd[0]").sendVKey(0)
        except pywintypes.com_error:
            logger.exception("Conexão COM SAP caiu!!!")
            console.print("COM SAP falhou, reconectando...", style="bold red")
            get_connection(token)
            time.sleep(10)
            while n_selected < count_connections():
                get_connection(token)
                time.sleep(10)
            session = create_session(n_selected)
            console.print("Reconectado com sucesso!!!", style="bold green")

        time.sleep(2)
