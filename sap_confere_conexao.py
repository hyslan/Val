'''Módulo para expor conexões e sessões ativas do SAP GUI'''
import win32com.client


def listar_conexoes():
    '''Função para listar as conexões ativas.'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connections = application.Children

    print("Conexões ativas:")
    for idx, connection in enumerate(connections):
        print(f"Conexão {idx}: {connection.Name}")


def listar_sessoes():
    '''Função para listar as sessions ativas'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)
    sessions = connection.Children

    print("Sessões ativas:")
    for idx in enumerate(sessions):
        # pylint: disable=W0212
        print(
            f"Sessão {idx}")
    return sessions


def criar_sessao(sessions):
    '''Função para criar sessões'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(0)

    # Obtendo o índice da última sessão ativa
    ultimo_indice = len(sessions) - 1
    print(len(sessions))

    # Criando uma nova sessão com base na última sessão ativa
    nova_sessao = connection.Children(ultimo_indice).CreateSession()
    while ultimo_indice >= len(sessions) - 1:
        sessions = connection.Children
        print(len(sessions))

    # Acessando a nova sessão
    session = connection.Children(len(sessions) - 1)
    return session
