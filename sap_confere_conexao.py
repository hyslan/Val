'''Módulo para expor conexões ativas do SAP GUI'''
import win32com.client


def listar_conexoes():
    '''Função para listar as conexões ativas.'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connections = application.Children

    print("Conexões ativas:")
    for idx, connection in enumerate(connections):
        print(f"Conexão {idx}: {connection.Name}")


def listar_sessoes(connection_index):
    '''Função para listar as sessions ativas'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    application = sapguiauto.GetScriptingEngine
    connection = application.Children(connection_index)
    sessions = connection.Children

    print(f"Sessões ativas na conexão {connection_index}:")
    for idx, session in enumerate(sessions):
        # pylint: disable=W0212
        print(
            f"Sessão {idx}: {session.Info.SystemName} - {session.Info._username_}")


# Exemplo de uso
listar_conexoes()
listar_sessoes(connection_index=0)
