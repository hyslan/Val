To use [[Multiple SapSession]] is needed to download as many as possible tx.sap from SAP Netweaver portal.

```
import win32com.client  
import pythoncom  
  
pythoncom.CoInitialize()  
sapguiauto = win32com.client.GetObject("SAPGUI")  
application = sapguiauto.GetScriptingEngine  
connection = application.Children(2)  
session: win32com.client.CDispatch = connection.Children(0)  
```
