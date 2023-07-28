# -Begin-----------------------------------------------------------------
'''Módulo para criar e gerar session do SAP GUI'''
# -Includes--------------------------------------------------------------
import win32com.client

# -Sub Main--------------------------------------------------------------


def conexao():
    '''Acessar session '''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    if not isinstance(sapguiauto, win32com.client.CDispatch):
        return

    application = sapguiauto.GetScriptingEngine
    if not isinstance(application, win32com.client.CDispatch):
        sapguiauto = None
        return

    application.HistoryEnabled = True

    connection = application.Children(0)
    if not isinstance(connection, win32com.client.CDispatch):
        application = None
        sapguiauto = None
        return

    if connection.DisabledByServer is True:
        connection = None
        application = None
        sapguiauto = None
        return

    return connection
