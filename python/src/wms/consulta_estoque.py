"""Módulo de consulta estoque de materiais"""
import time
import os
import xlwings as xw
import pandas as pd
from python.src import sap


def estoque(session, contrato, n_con):
    """Função para consultar estoque"""
    caminho = os.getcwd() + "\\sheets\\"
    session.StartTransaction("MBLB")
    frame = session.findById("wnd[0]")
    frame.findByid("wnd[0]/usr/ctxtLIFNR-LOW").text = contrato
    print("Consultando Estoque de Materiais")
    frame.SendVkey(8)
    frame.sendVKey(42)  # Lista Detalhada
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.contextMenu()
    grid.selectContextMenuItem("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById(
        "wnd[1]/usr/ctxtDY_PATH").text = caminho
    session.findById(
        "wnd[1]/usr/ctxtDY_FILENAME").text = f"estoque_{contrato}.XLSX"
    session.findById("wnd[1]").sendVKey(11)  # Substituir
    print("Planilha de estoque gerada com sucesso.")
    materiais = pd.read_excel(caminho + f"estoque_{contrato}.XLSX",
                              sheet_name="Sheet1", usecols=["Material",
                                                            "Texto breve material",
                                                            "Utilização livre"
                                                            ]
                              )
    materiais = materiais.dropna()
    materiais['Material'] = materiais['Material'].astype(int).astype(str)
    sessao = sap
    con = sessao.connection_object(n_con)
    total_sessoes = sessao.contar_sessoes(n_con)
    if not total_sessoes == 6:
        try:
            print("Encerrando Sessão.")
            con.CloseSession(f"/app/con[0]/ses[{total_sessoes - 1}]")
            print("Sessão Encerrada.")
        except Exception as e:
            print(f"Erro em consulta_estoque - Encerrar Sessão:{e}")
        time.sleep(3)

    print("Fechando Arquivo Excel.\n")
    try:
        time.sleep(8)
        book = xw.Book(f'estoque_{contrato}.xlsx')
        book.app.quit()
    except Exception as e:
        print(f"Erro em consulta_estoque - MS EXCEL:{e}")

    return materiais