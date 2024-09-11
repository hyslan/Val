# testa_material_sap.py
"""Módulo verifica GRID SAP de materiais se está vazio."""

from __future__ import annotations

import logging
import typing

import pywintypes

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch

logger = logging.getLogger(__name__)


def testa_material_sap(tb_materiais: CDispatch) -> str | None:
    """Módulo de verificar materiais inclusos na ordem."""
    try:
        cell = tb_materiais.GetCellValue(0, "MATERIAL")
        return str(cell) if cell is not None else None
    except pywintypes.com_error:
        logger.info("Não foi possível acessar o GRID de materiais.")
        return None
