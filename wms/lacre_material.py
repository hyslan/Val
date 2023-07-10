'''Módulo Lacre SABESP.'''
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets

session = connect_to_sap()
(
    lista,
    _,
    _,
    _,
    planilha,
    _,
    _,
    _,
    _,
    _,
    tb_contratada,
    _,
    *_,
) = load_worksheets()


def caca_lacre(tb_materiais, etapa):
    '''Módulo de procurar lacres no grid de materiais.'''
    num_material_linhas = tb_materiais.RowCount  # Conta as Rows
    # Número da Row do Grid Materiais do SAP
    n_material = 0
    procura_lacre = []
    ultima_linha_material = num_material_linhas
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(
            n_material, "MATERIAL")
        procura_lacre.append(sap_material)
    if '50001070' not in procura_lacre:

        tb_materiais.InsertRows(str(ultima_linha_material))
        tb_materiais.modifyCell(
            ultima_linha_material, "ETAPA", etapa
        )
        tb_materiais.modifyCell(
            ultima_linha_material, "MATERIAL", "50001070"
        )
        tb_materiais.modifyCell(
            ultima_linha_material, "QUANT", "1"
        )
        tb_materiais.setCurrentCell(
            ultima_linha_material, "QUANT"
        )
        ultima_linha_material = ultima_linha_material + 1
