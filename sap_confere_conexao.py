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
    for idx, session in enumerate(sessions):
        # pylint: disable=W0212
        print(
            f"Sessão {idx}: {session.Info.SystemName} - {session.Info._username_}")
