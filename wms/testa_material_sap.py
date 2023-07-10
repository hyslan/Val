# testa_material_sap.py
'''Módulo verifica GRID SAP de materiais se está vazio.'''
import pywintypes
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


def testa_material_sap(int_num_lordem, tb_materiais):
    '''Módulo de verificar materiais inclusos na ordem.'''
    try:
        sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
        print("Tem material vinculado.")
        return sap_material
    # pylint: disable=E1101
    except pywintypes.com_error:
        material_obs = planilha.cell(row=int_num_lordem, column=3)
        material_obs.value = "Sem Material Vinculado"
        print("Sem material vinculado.")
        lista.save('lista.xlsx')  # salva Planilha
        return None
