"""Módulo Lacre SABESP."""
from src.wms.localiza_material import btn_busca_material


def caca_lacre(tb_materiais, etapa, estoque, session):
    """Módulo de procurar lacres no grid de materiais."""
    num_material_linhas = tb_materiais.RowCount  # Conta as Rows
    lacre_estoque = estoque[estoque['Material'] == '50001070']
    procura_lacre = []
    ultima_linha_material = num_material_linhas
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(
            n_material, "MATERIAL")
        procura_lacre.append(sap_material)

    if '50001070' in procura_lacre:
        btn_busca_material(tb_materiais, session, '50001070')
        quantidade = tb_materiais.GetCellValue(
            tb_materiais.CurrentCellRow, "QUANT"
        )
        qtd_float = float(quantidade.replace(",", "."))
        if qtd_float >= 2.00 and not lacre_estoque.empty:
            tb_materiais.modifyCell(
                tb_materiais.CurrentCellRow, "QUANT", "1"
            )
            tb_materiais.setCurrentCell(
                tb_materiais.CurrentCellRow, "QUANT"
            )

        # Caso lacre sem estoque
        if lacre_estoque.empty:
            try:
                tb_materiais.modifyCell(
                    tb_materiais.CurrentCellRow, "QUANT", "0"
                )
                tb_materiais.setCurrentCell(
                    tb_materiais.CurrentCellRow, "QUANT"
                )
                tb_materiais.modifyCheckBox(
                    tb_materiais.CurrentCellRow, "ELIMINADO", True
                )
            except:
                pass

    if '50001070' not in procura_lacre and not lacre_estoque.empty:

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
