import os
import win32com.client
import pythoncom
import platform
# -Sub Main--------------------------------------------------------------

'''Função para conexão SAP'''
pythoncom.CoInitialize()
sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")