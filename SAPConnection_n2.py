#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
import sys
import win32com.client

#-Sub Main--------------------------------------------------------------
def Connect_to_SAP_n2():

  

    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    if not type(SapGuiAuto) == win32com.client.CDispatch:
      return

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
      SapGuiAuto = None
      return

    application.HistoryEnabled = True

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
      application = None
      SapGuiAuto = None
      return

    if connection.DisabledByServer == True:
      connection = None
      application = None
      SapGuiAuto = None
      return

    session_n2 = connection.Children(1)
    if not type(session_n2) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      return

    if session_n2.Busy == True:
      session_n2 = None
      connection = None
      application = None
      SapGuiAuto = None
      return

    if session_n2.Info.IsLowSpeedConnection == True:
      session_n2 = None
      connection = None
      application = None
      SapGuiAuto = None
      return

    return session_n2

