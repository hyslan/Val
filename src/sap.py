"""Módulo para interagir com o SAP GUI"""
import os
import subprocess
import threading
import time
from arrow import get
import win32com.client
import pythoncom
from rich.console import Console
import rich.console
from dotenv import load_dotenv

console: rich.console.Console = Console()


def reconnect(session):
    console.print("Tentando obter de volta IID_IDispatch")
    s_id = pythoncom.CoMarshalInterThreadInterfaceInStream(
        pythoncom.IID_IDispatch, session)
    console.print(f"Obtido com sucesso: {s_id}")
    return s_id


def connection_object(n_selected) -> win32com.client.CDispatch:
    """Função para listar as conexões ativas."""
    con = get_app()
    connection: win32com.client.CDispatch = con.Item(n_selected)
    return connection


def listar_sessoes(n_selected) -> win32com.client.CDispatch:
    """Função para listar as sessions ativas"""
    con = get_app()
    con_selected: win32com.client.CDispatch = con.Item(n_selected)
    sessions: win32com.client.CDispatch = con_selected.Sessions
    return sessions


def contar_sessoes(n_selected) -> int:
    """Contar por tamanho de 1 a 6, caso for criar sessão subtrair -1"""
    con = get_app()
    con_selected: win32com.client.CDispatch = con.Item(n_selected)
    sessions: win32com.client.CDispatch = con_selected.Sessions
    console.print(
        f"[blue italic]Quantidade de sessões ativas: {sessions.Count}")
    return sessions.Count


def create_session(n_selected: int) -> win32com.client.CDispatch:
    """Função para criar sessões"""
    con = get_app()
    con_selected: win32com.client.CDispatch = con.Item(n_selected)
    sessions = con_selected.Sessions
    # Obtendo o índice da última sessão ativa
    ultimo_indice = len(sessions) - 1

    # Criando uma nova sessão com base na última sessão ativa
    if ultimo_indice < 5:
        con_selected.Children(ultimo_indice).CreateSession()
        while ultimo_indice >= len(sessions) - 1:
            sessions = con_selected.Children

        # Acessando a nova sessão
        session = con_selected.Children(len(sessions) - 1)
    else:
        session = con_selected.Children(5)

    return session


def choose_connection(n_selected: int) -> win32com.client.CDispatch:
    """Escolher com qual sessão trabalhar"""
    con = get_app()
    session: win32com.client.CDispatch = con.Item(n_selected).Sessions(0)
    return session


def fechar_conexao(n_con) -> None:
    """Função para fechar o SAP."""
    con = get_app()
    connection = con.Item(n_con)
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


def get_connection(token: str) -> str:
    load_dotenv()
    sap_access = (
        '[System]\n'
        'Name=EP0\n'
        'Client=100\n'
        fr'GuiParm={os.environ['SERVER']}'
        '\n'
        '[User]\n'
        f'Name={os.environ['USR']}\n'
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
    path_archive = os.getcwd() + '\\shortcut\\repeat\\tx.sap'
    print("Saving the SAP access file...")
    with open(path_archive, 'w') as s:
        s.write(sap_access)

    # Execute the command
    try:
        subprocess.run(["powershell", "start", '"' + path_archive + '"'],
                       shell=True, check=False)
        time.sleep(5)
        # Verifica se o processo está em execução
        if not is_process_running("powershell.exe"):
            print("Erro: O arquivo não foi aberto corretamente.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao iniciar o arquivo tx.sap: {e}")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    return token


def is_process_running(process_name: str):
    """Verifica se o processo está em execução"""
    try:
        subprocess.check_output(
            f'tasklist /FI "IMAGENAME eq {process_name}"', shell=True)
        return True
    except subprocess.CalledProcessError:
        return False


def get_app() -> win32com.client.CDispatch:
    """Get the SAP GUI application

    Returns:
        win32com.client.CDispatch: Connections from application
    """
    pythoncom.CoInitialize()
    app: win32com.client.CDispatch = win32com.client.GetObject(
        "SAPGUI").GetScriptingEngine
    con = app.Connections
    return con
