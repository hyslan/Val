#-Begin-----------------------------------------------------------------

#-Includes--------------------------------------------------------------
import sys
import win32com.client

#-Sub Main--------------------------------------------------------------
def Connect_to_SAP():

  

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

    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
      connection = None
      application = None
      SapGuiAuto = None
      return

    if session.Busy == True:
      session = None
      connection = None
      application = None
      SapGuiAuto = None
      return

    if session.Info.IsLowSpeedConnection == True:
      session = None
      connection = None
      application = None
      SapGuiAuto = None
      return

    return session

