'''Módulo de consulta estoque de materiais'''
import time
import xlwings as xw
import pandas as pd
from src import sap


def estoque(session, sessions, contrato):
    '''Função para consultar estoque'''
    caminho = "sheets/"
    session.StartTransaction("MBLB")
    frame = session.findById("wnd[0]")
    frame.findByid("wnd[0]/usr/ctxtLIFNR-LOW").text = contrato[0]
    print("Consultando Estoque de Materiais")
    frame.SendVkey(8)
    frame.sendVKey(42)  # Lista Detalhada
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById(
        "wnd[1]/usr/ctxtDY_PATH").text = caminho
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "estoque.XLSX"
    session.findById("wnd[1]").sendVKey(11)  # Substituir
    print("Planilha de estoque gerada com sucesso.")
    materiais = pd.read_excel(caminho + "estoque.XLSX",
                              sheet_name="Sheet1", usecols=["Material",
                                                            "Texto breve material",
                                                            "Utilização livre"
                                                            ]
                              )
    materiais = materiais.dropna()
    materiais['Material'] = materiais['Material'].astype(int).astype(str)
    sessao = sap.Sap()
    con = sessao.listar_conexoes()
    total_sessoes = sessao.contar_sessoes()
    if not total_sessoes == 6:
        # session.EndTransaction()
        print("Encerrando Sessão.")
        con.CloseSession(f"/app/con[0]/ses[{len(sessions)}]")
        time.sleep(6)

    print("Fechando Arquivo Excel.\n")
    try:
        book = xw.Book('estoque.xlsx')
        book.close()
    except Exception as e:
        print(e)

    return materiais
