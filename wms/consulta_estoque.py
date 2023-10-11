'''Módulo de consulta estoque de materiais'''
import sap


def estoque_novasp(session, sessions):
    '''Função para consultar estoque NOVASP'''
    estoque = set()
    session.StartTransaction("MBLB")
    frame = session.findById("wnd[0]")
    frame.findByid("wnd[0]/usr/ctxtLIFNR-LOW").text = "4600041302"
    print("Consultando Estoque NOVASP.")
    frame.SendVkey(8)
    frame.sendVKey(42)  # Lista Detalhada
    grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    grid.selectColumn("MATNR")
    grid.currentCellRow = 0
    total = grid.RowCount
    print("Obtendo os materiais disponiveis.")
    for i in range(0, total):
        material = grid.GetCellValue(i, "MATNR")
        estoque.add(material)

    print("Consulta concluída.")
    con = sap.listar_conexoes()
    # session.EndTransaction()
    con.CloseSession(f"/app/con[0]/ses[{len(sessions)}]")

    return estoque
