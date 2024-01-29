'''Módulo para interagir com o SAP GUI'''
import subprocess
import win32com.client
import pythoncom
from rich.console import Console
from rich.prompt import Prompt

console = Console()


class Sap():
    '''Classe para modelagem de métodos na 
    manipulação do COM SapGui'''

    def __init__(self) -> None:
        pass

    def listar_conexoes(self):
        '''Função para listar as conexões ativas.'''
        # pylint: disable=E1101
        pythoncom.CoInitialize()
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connections = application.Children(0)

        return connections

    def listar_sessoes(self):
        '''Função para listar as sessions ativas'''
        # pylint: disable=E1101
        pythoncom.CoInitialize()
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        sessions = connection.Children

        return sessions

    def contar_sessoes(self):
        '''Contar por tamanho de 1 a 6, caso for criar sessão subtrair -1'''
        # pylint: disable=E1101
        pythoncom.CoInitialize()
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        sessions = connection.Children
        console.print(
            f"[blue italic]Quantidade de sessões ativas: {sessions.Count}")
        return sessions.Count

    def criar_sessao(self, sessions):
        '''Função para criar sessões'''
        # pylint: disable=E1101
        pythoncom.CoInitialize()
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)

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

    def escolher_sessao(self):
        '''Escolher com qual sessão trabalhar'''
        total = self.contar_sessoes()
        n_selected = Prompt.ask(
            "[bold]Escolha entre as sessões disponíveis, lembrando que sessão 1 é nº 0",
            choices=list(range(total)), default=0)
        conn = self.listar_sessoes()
        session = conn.Children(n_selected)
        return session

    def fechar_conexao(self):
        '''Função para fechar o SAP.'''
        # pylint: disable=E1101
        pythoncom.CoInitialize()
        sapguiauto = win32com.client.GetObject("SAPGUI")
        application = sapguiauto.GetScriptingEngine
        connection = application.Children(0)
        connection.CloseConnection()

    def encerrar_sap(self):
        '''Encerra o app SAP'''
        processo = 'saplogon.exe'
        try:
            subprocess.run(['taskkill', '/F', '/IM', processo], check=True)
            print(f'O processo {processo} foi encerrado com sucesso.')
        except subprocess.CalledProcessError:
            print(f'Não foi possível encerrar o processo {processo}.')
