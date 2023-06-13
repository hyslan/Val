# sap_connection.py
# -Begin-----------------------------------------------------------------
'''Módulo SAP'''
# -Bibliotecas--------------------------------------------------------------
import win32com.client

# -Sub Main--------------------------------------------------------------


def connect_to_sap():
    '''Função para conexão SAP'''
    sapguiauto = win32com.client.GetObject("SAPGUI")
    # sapguiauto = win32com.client.GetObject("SAPGUI")
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

    session = connection.Children(0)
    if not isinstance(session, win32com.client.CDispatch):
        connection = None
        application = None
        sapguiauto = None
        return

    if session.Busy is True:
        session = None
        connection = None
        application = None
        sapguiauto = None
        return

    if session.Info.IsLowSpeedConnection is True:
        session = None
        connection = None
        application = None
        sapguiauto = None
        return

    return session
