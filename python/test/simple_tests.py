import contextlib

import pythoncom
import win32com.client

# -Sub Main--------------------------------------------------------------

"""Função para conexão SAP"""
# pythoncom.CoInitialize()
# sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")

with contextlib.suppress(pythoncom.com_error):
    sapguiauto = win32com.client.GetObject(
        "SapROTWR.SapROTWrapper").GetROTEntry("SAPGUI")
