'''Módulo para interagir com o SAP GUI'''
import subprocess
import win32com.client


def listar_conexoes():
    '''Função para listar as conexões ativas.'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connections = application.Children(0)

    return connections


def listar_sessoes():
    '''Função para listar as sessions ativas'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    sessions = connection.Children

    return sessions


def criar_sessao(sessions):
    '''Função para criar sessões'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)

    # Obtendo o índice da última sessão ativa
    ultimo_indice = len(sessions) - 1

    # Criando uma nova sessão com base na última sessão ativa
    connection.Children(ultimo_indice).CreateSession()
    while ultimo_indice >= len(sessions) - 1:
        sessions = connection.Children

    # Acessando a nova sessão
    session = connection.Children(len(sessions) - 1)
    return session


def fechar_conexao():
    '''Função para fechar o SAP.'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    connection.CloseConnection()


def encerrar_sap():
    '''Encerra o app SAP'''
    processo = 'saplogon.exe'
    try:
        subprocess.run(['taskkill', '/F', '/IM', processo], check=True)
        print(f'O processo {processo} foi encerrado com sucesso.')
    except subprocess.CalledProcessError:
        print(f'Não foi possível encerrar o processo {processo}.')
