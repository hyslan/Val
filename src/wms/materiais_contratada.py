"""Módulo de materiais contratada"""
from typing import List
import pywintypes
from src.excel_tbs import load_worksheets
from src.wms import lacre_material

(
    _,
    _,
    _,
    _,
    _,
    _,
    _,
    _,
    _,
    _,
    tb_contratada,
    tb_contratada_gb,
    _,
    *_,
) = load_worksheets()


def materiais_novasp(tb_materiais,
                     num_material_linhas,
                     lacre,
                     estoque,
                     session):
    """Contratada NOVASP - MLG"""
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(
            n_material, "MATERIAL")
        sap_etapa_material = tb_materiais.GetCellValue(
            n_material, "ETAPA")
        material_estoque = estoque[estoque['Material'] == sap_material]

        # Verifica se está na lista tb_contratada
        if sap_material in tb_contratada:
            # Marca Contratada
            tb_materiais.modifyCheckbox(
                n_material, "CONTRATADA", True)
            continue

        if lacre is False:
            if sap_material in ('50000328', '50000263'):
                # Remove o lacre bege antigo.
                tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
                lacre_material.caca_lacre(
                    tb_materiais, sap_etapa_material,
                    estoque, session)
                lacre = True
            if sap_material == '50001070':
                lacre_material.caca_lacre(
                    tb_materiais, sap_etapa_material,
                    estoque, session)
                lacre = True

        if sap_material == '10014780':
            # Remove REPARADOR ASFALTO MOD FX C DER/IV PMSP
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )

        # --- APENAS SE ENCARREGADOS PEDIREM ---
        # try:
        #     if sap_material == '10014709':
        #         # Marca Contratada
        #         tb_materiais.modifyCheckbox(
        #             n_material, "CONTRATADA", True)
        #         print("Aslfato frio da NOVASP por enquanto.")
        # # pylint: disable=E1101
        # except pywintypes.com_error:
        #     print(
        #         f"Etapa: {sap_etapa_material} - Asfalto frio já foi retirado.")

        # try:
        #     if sap_material == '30028856':
        #         # Marca Contratada
        #         tb_materiais.modifyCheckbox(
        #             n_material, "CONTRATADA", True)
        #         print("TUBO ESG DN 100 da NOVASP por enquanto.")
        # # pylint: disable=E1101
        # except pywintypes.com_error:
        #     print(
        #         f"Etapa: {sap_etapa_material} - TUBO ESG DN 100 já foi retirado.")

        # try:
        #     if sap_material == '30026319':
        #         # Marca Contratada
        #         tb_materiais.modifyCheckbox(
        #             n_material, "CONTRATADA", True)
        #         print("TUBO PVC RIG PB JEI/JERI DN 100 da NOVASP por enquanto.")
        # # pylint: disable=E1101
        # except pywintypes.com_error:
        #     print(
        #         f"Etapa: {sap_etapa_material} - TUBO PVC RIG PB JEI/JERI DN 100 já foi retirado.")

        try:
            if sap_material == '30001865' and material_estoque.empty:
                tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True)
                # Adiciona União PEAD vigente.
                check_uniao = estoque[estoque['Material'] == '30029526']
                if not check_uniao.empty:
                    try:
                        tb_materiais.modifyCell(
                            ultima_linha_material, "MATERIAL", "30029526"
                        )
                        tb_materiais.modifyCell(
                            ultima_linha_material, "QUANT", "1"
                        )
                        tb_materiais.setCurrentCell(
                            ultima_linha_material, "QUANT"
                        )
                        ultima_linha_material = ultima_linha_material + 1
                    except pywintypes.com_error:
                        print(
                            "Erro ao adicionar União PEAD vigente.")
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(
                f"Etapa: {sap_etapa_material} - UNIAO P/TUBO PEAD DE 20 MM já foi retirado.")


def materiais_gb_itaquera(tb_materiais,
                          num_material_linhas,
                          lacre,
                          ultima_linha_material,
                          estoque,
                          session):
    """Contratada GB - MLN"""
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(
            n_material, "MATERIAL")
        sap_etapa_material = tb_materiais.GetCellValue(
            n_material, "ETAPA")

        # Verifica se está na lista tb_contratada
        if sap_material in tb_contratada_gb:
            # Marca Contratada
            tb_materiais.modifyCheckbox(
                n_material, "CONTRATADA", True)
            continue

        if lacre is False:
            if sap_material in ('50000328', '50000263'):
                # Remove o lacre bege antigo.
                tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
                lacre_material.caca_lacre(
                    tb_materiais, sap_etapa_material,
                    estoque, session)
                lacre = True

        if sap_material == '10014780':
            # Remove REPARADOR ASFALTO MOD FX C DER/IV PMSP
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )

        # if sap_material == '50000178':
        #     # Remove DISPOSITIVO MED PLASTICO DN 20
        #     tb_materiais.modifyCheckbox(
        #         n_material, "ELIMINADO", True
        #     )
        #     tb_materiais.InsertRows(str(ultima_linha_material))
        #     tb_materiais.modifyCell(
        #         ultima_linha_material, "ETAPA", sap_etapa_material
        #     )
        #     # Adiciona DISPOSITIVO MED METAL DN 20.
        #     tb_materiais.modifyCell(
        #         ultima_linha_material, "MATERIAL", "50000021"
        #     )
        #     tb_materiais.modifyCell(
        #         ultima_linha_material, "QUANT", "1"
        #     )
        #     tb_materiais.setCurrentCell(
        #         ultima_linha_material, "QUANT"
        #     )
        #     ultima_linha_material = ultima_linha_material + 1

        # try:
        #     if sap_material == '10014709':
        #         # Marca Contratada
        #         tb_materiais.modifyCheckbox(
        #             n_material, "CONTRATADA", True)
        #         print("Aslfato frio da NOVASP por enquanto.")
        # # pylint: disable=E1101
        # except pywintypes.com_error:
        #     print(
        #         f"Etapa: {sap_etapa_material} - Asfalto frio já foi retirado.")

        # try:
        #     if sap_material == '30028856':
        #         # Marca Contratada
        #         tb_materiais.modifyCheckbox(
        #             n_material, "CONTRATADA", True)
        #         print("TUBO ESG DN 100 da NOVASP por enquanto.")
        # # pylint: disable=E1101
        # except pywintypes.com_error:
        #     print(
        #         f"Etapa: {sap_etapa_material} - TUBO ESG DN 100 já foi retirado.")

        # try:
        #     if sap_material == '30026319':
        #         # Marca Contratada
        #         tb_materiais.modifyCheckbox(
        #             n_material, "CONTRATADA", True)
        #         print("TUBO PVC RIG PB JEI/JERI DN 100 da NOVASP por enquanto.")
        # # pylint: disable=E1101
        # except pywintypes.com_error:
        #     print(
        #         f"Etapa: {sap_etapa_material} - TUBO PVC RIG PB JEI/JERI DN 100 já foi retirado.")

        try:
            if sap_material in ('30001865', '30000882'):
                # Remove UNIAO P/TUBO PEAD DE 20 MM.
                tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
                tb_materiais.InsertRows(str(ultima_linha_material))
                tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", sap_etapa_material
                )
                # Adiciona União PEAD vigente.
                tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "30029526"
                )
                tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1"
                )
                tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT"
                )
                ultima_linha_material = ultima_linha_material + 1
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(
                f"Etapa: {sap_etapa_material} - UNIAO P/TUBO PEAD DE 20 MM já foi retirado.")


def materiais_contratada(tb_materiais, contrato, estoque, session):
    """Módulo de materiais da NOVASP."""
    num_material_linhas = tb_materiais.RowCount  # Conta as Rows
    # Número da Row do Grid Materiais do SAP
    lacre = False
    ultima_linha_material = num_material_linhas
    match contrato:
        case "4600041302":
            materiais_novasp(tb_materiais,
                             num_material_linhas,
                             lacre,
                             estoque,
                             session)
        case "4600042975":
            materiais_novasp(tb_materiais,
                             num_material_linhas,
                             lacre,
                             estoque,
                             session)
        case "4600042888":
            materiais_gb_itaquera(tb_materiais,
                                  num_material_linhas,
                                  lacre,
                                  ultima_linha_material,
                                  estoque,
                                  session)
        case "4600056089":
            materiais_gb_itaquera(tb_materiais,
                                  num_material_linhas,
                                  lacre,
                                  ultima_linha_material,
                                  estoque,
                                  session)
        case _:
            return


def lista_materiais() -> List[str]:
    """Retorna lista de materiais marcados como contratada."""
    contratada_list = [
        '50000017', '50000333', '30006741', '30002033',
        '30001023', '30004491', '30000732', '30001363',
        '50000318', '30005788', '30002832', '30002842',
        '30003288', '50000179', '10006345', '10006348',
        '50000180', '30003520', '30013576', '30003675',
        '30003758', '30007217', '50000307', '50000150',
        '50000076', '30000880', '30007737', '30012631',
    ]
    return contratada_list
