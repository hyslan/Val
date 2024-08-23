# testa_material_sap.py
"""Módulo verifica GRID SAP de materiais se está vazio."""
import pywintypes


def testa_material_sap(tb_materiais):
    """Módulo de verificar materiais inclusos na ordem."""
    try:
        return tb_materiais.GetCellValue(0, "MATERIAL")
    # pylint: disable=E1101
    except pywintypes.com_error:
        return None
