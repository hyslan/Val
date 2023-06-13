from sap_connection import connect_to_sap
session = connect_to_sap()

tb_materiais = session.findById(
    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM/ssubSUB_TAB:"
    + "ZSBMM_VALORACAOINV:9030/cntlCC_MATERIAIS/shellcont/shell"
)
sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
if sap_material is not None:  # Se existem materiais, irá contar
    num_material_linhas = tb_materiais.RowCount  # Conta as Rows
    print(f"Qtd de linhas de materiais: {num_material_linhas}")
    # Número da Row do Grid Materiais do SAP
    n_material = 0
    ultima_linha_material = num_material_linhas
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(n_material, "MATERIAL")
        sap_etapa_material = tb_materiais.GetCellValue(
            n_material, "ETAPA")
        # Verifica se está na lista tb_contratada
        if sap_material == '50000328':
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )
            tb_materiais.InsertRows(str(ultima_linha_material))
            tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", sap_etapa_material
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
