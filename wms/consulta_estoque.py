'''Módulo de consulta estoque de materiais'''
import time
import pandas as pd
import sap_sessoes


def estoque_novasp():
    '''Função para consultar estoque NOVASP'''
    connection = sap_sessoes.conexao()
    sessions = connection.Children
    n_sessao = 0
    estoque = set()

    for sessao in sessions:
        print(f"Sessões: {sessao.Info.SessionNumber}")
        n_sessao += 1

    if n_sessao < 6:
        sessao = connection.Children(n_sessao - 1)
        sessao.CreateSession()
        print("Sessão Criada")

    # Espera até que a nova sessão esteja pronta (timeout de 60 segundos)
    timeout = 60
    start_time = time.time()
    while time.time() - start_time < timeout:
        if len(connection.Children) == n_sessao + 1:
            print("Nova sessão aberta!")
            sessao_nova = connection.Children(n_sessao)
            break
        time.sleep(1)
    else:
        print("Timeout: Nova sessão não foi aberta.")

    sessao_nova.StartTransaction("MBLB")
    frame = sessao_nova.findById("wnd[0]")
    frame.findByid("wnd[0]/usr/ctxtLIFNR-LOW").text = "4600041302"
    print("Consultando Estoque NOVASP.")
    frame.SendVkey(8)
    frame.sendVKey(42)  # Lista Detalhada
    grid = sessao_nova.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.selectColumn("MATNR")
    grid.currentCellRow = 0
    total = grid.RowCount
    print("Obtendo os materiais disponiveis.")
    for i in range(0, total):
        material = grid.GetCellValue(i, "MATNR")
        estoque.add(material)

    print("Consulta concluída.")
    sessao_nova.CloseSession(n_sessao)

    return estoque
