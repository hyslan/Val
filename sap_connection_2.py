# -Begin-----------------------------------------------------------------
'''Módulo para session 2 SAP GUI'''
# -Includes--------------------------------------------------------------
import win32com.client

# -Sub Main--------------------------------------------------------------


def connect_to_sap_2():
    '''Acessar session 2'''
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

    session2 = connection.Children(1)
    if not isinstance(session2, win32com.client.CDispatch):
        connection = None
        application = None
        sapguiauto = None
        return

    if session2.Busy is True:
        session2 = None
        connection = None
        application = None
        sapguiauto = None
        return

    if session2.Info.IsLowSpeedConnection is True:
        session2 = None
        connection = None
        application = None
        sapguiauto = None
        return

    return session2
