'''Módulo de materiais contratada'''
import pywintypes
from excel_tbs import load_worksheets
from wms import lacre_material

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


def materiais_novasp(tb_materiais,
                     num_material_linhas,
                     n_material,
                     lacre,
                     ultima_linha_material):
    '''Contratada NOVASP - MLG'''
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(
            n_material, "MATERIAL")
        sap_etapa_material = tb_materiais.GetCellValue(
            n_material, "ETAPA")

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
                lacre_material.caca_lacre(tb_materiais, sap_etapa_material)
                lacre = True

        if sap_material in ('50000328', '50000263'):
            # Remove o lacre bege antigo.
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )

        if sap_material == '10014780':
            # Remove REPARADOR ASFALTO MOD FX C DER/IV PMSP
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )

        try:
            if sap_material == '10014709':
                # Marca Contratada
                tb_materiais.modifyCheckbox(
                    n_material, "CONTRATADA", True)
                print("Aslfato frio da NOVASP por enquanto.")
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(
                f"Etapa: {sap_etapa_material} - Asfalto frio já foi retirado.")

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
            if sap_material == '30001865':
                # Marca Contratada
                tb_materiais.modifyCheckbox(
                    n_material, "CONTRATADA", True)
                print("UNIAO P/TUBO PEAD DE 20 MM da NOVASP por enquanto.")
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(
                f"Etapa: {sap_etapa_material} - UNIAO P/TUBO PEAD DE 20 MM já foi retirado.")


def materiais_gb_itaquera(tb_materiais,
                          num_material_linhas,
                          n_material,
                          lacre,
                          ultima_linha_material):
    '''Contratada GB - MLN'''
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(
            n_material, "MATERIAL")
        sap_etapa_material = tb_materiais.GetCellValue(
            n_material, "ETAPA")

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
                lacre_material.caca_lacre(tb_materiais, sap_etapa_material)
                lacre = True

        if sap_material in ('50000328', '50000263'):
            # Remove o lacre bege antigo.
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )

        if sap_material == '10014780':
            # Remove REPARADOR ASFALTO MOD FX C DER/IV PMSP
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )

        if sap_material == '50000178':
            # Remove DISPOSITIVO MED PLASTICO DN 20
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )
            tb_materiais.InsertRows(str(ultima_linha_material))
            tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", sap_etapa_material
            )
            # Adiciona DISPOSITIVO MED METAL DN 20.
            tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "50000021"
            )
            tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1

        try:
            if sap_material == '10014709':
                # Marca Contratada
                tb_materiais.modifyCheckbox(
                    n_material, "CONTRATADA", True)
                print("Aslfato frio da NOVASP por enquanto.")
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(
                f"Etapa: {sap_etapa_material} - Asfalto frio já foi retirado.")

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
            if sap_material == '30001865':
                # Marca Contratada
                tb_materiais.modifyCheckbox(
                    n_material, "CONTRATADA", True)
                print("UNIAO P/TUBO PEAD DE 20 MM da NOVASP por enquanto.")
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(
                f"Etapa: {sap_etapa_material} - UNIAO P/TUBO PEAD DE 20 MM já foi retirado.")


def materiais_contratada(tb_materiais, contrato):
    '''Módulo de materiais da NOVASP.'''
    num_material_linhas = tb_materiais.RowCount  # Conta as Rows
    # Número da Row do Grid Materiais do SAP
    n_material = 0
    lacre = False
    ultima_linha_material = num_material_linhas
    match contrato:
        case "4600041302":
            materiais_novasp(tb_materiais,
                             num_material_linhas,
                             n_material,
                             lacre,
                             ultima_linha_material)
        case "4600042888":
            materiais_gb_itaquera(tb_materiais,
                                  num_material_linhas,
                                  n_material,
                                  lacre,
                                  ultima_linha_material)
        case _:
            return
