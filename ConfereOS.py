# ConfereOS.py
import sys
import win32com.client


def ConsultaOS(n_os):
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.StartTransaction("ZSBPM020")
    StatusUsuario = "USTXT"
    StatusSistema = "STTXT"
    Campo_OS = session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW")
    Campo_OS.Text = (n_os)
    session.findById("wnd[0]/usr/txtS_CONTR-LOW").text = "4600041302"
    session.findById("wnd[0]/usr/txtS_UN_ADM-LOW").text = "344"
    session.findById("wnd[0]").sendVKey(8)
    Consulta = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    StatusSistema = Consulta.GetCellValue(0, "STTXT")
    StatusUsuario = Consulta.GetCellValue(0, "USTXT")
    Corte = Consulta.GetCellValue(0, "ZZLOCAL_CORTE")  # Supressão
    Relig = Consulta.GetCellValue(0, "ZZLOCAL_RELIGA")  # Religação
    PosicaoRede = Consulta.GetCellValue(
        0, "ZZPOSICAO_REDE")  # Posição da Rede
    Profundidade = Consulta.GetCellValue(0, "ZZPROFUNDIDADE")
    session.findById("wnd[0]").sendVKey(3)  # Voltar
    return StatusSistema, StatusUsuario, Corte, Relig, PosicaoRede, Profundidade
