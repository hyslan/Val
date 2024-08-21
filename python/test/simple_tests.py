import pythoncom
import win32com.client

# -Sub Main--------------------------------------------------------------

"""Função para conexão SAP"""
# pythoncom.CoInitialize()
# sapguiauto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")

try:
    sapguiauto = win32com.client.GetObject(
        "SapROTWR.SapROTWrapper").GetROTEntry("SAPGUI")
    print("Objeto SAPGUI obtido com sucesso")
except pythoncom.com_error as e:
    print(f"Erro ao obter o objeto SAPGUI: {e}")
