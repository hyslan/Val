# testa_material_sap.py
'''Módulo verifica GRID SAP de materiais se está vazio.'''
import pywintypes


def testa_material_sap(tb_materiais):
    '''Módulo de verificar materiais inclusos na ordem.'''
    try:
        sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
        print("Tem material vinculado.")
        return sap_material
    # pylint: disable=E1101
    except pywintypes.com_error:
        print("Sem material vinculado.")
        return None
