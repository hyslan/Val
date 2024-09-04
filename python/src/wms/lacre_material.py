"""Módulo Lacre SABESP."""

from __future__ import annotations

import logging
import typing

import pywintypes

from python.src.wms.localiza_material import btn_busca_material

if typing.TYPE_CHECKING:
    from pandas import DataFrame
    from win32com.client import CDispatch

logger = logging.getLogger(__name__)


def busca_material_etapa(procura_lacre: list[dict[str, str]], material: str, etapa: str) -> bool:
    """Verifica se o material e a etapa estão presentes em procura_lacre."""
    return any(item["Material"] == material and item["Etapa"] == etapa for item in procura_lacre)


def caca_lacre(tb_materiais: CDispatch, etapa: str, estoque: DataFrame, session: CDispatch) -> None:
    """Módulo de procurar lacres no grid de materiais."""
    num_material_linhas = tb_materiais.RowCount  # Conta as Rows
    lacre_estoque = estoque[estoque["Material"] == "50001070"]
    procura_lacre = []
    ultima_linha_material = num_material_linhas
    max_qtd = 2.00
    # Loop do Grid Materiais.
    for n_material in range(num_material_linhas):
        # Pega valor da célula 0
        sap_material = tb_materiais.GetCellValue(n_material, "MATERIAL")
        sap_etapa_material = tb_materiais.GetCellValue(n_material, "ETAPA")
        procura_lacre.append({"Material": sap_material, "Etapa": sap_etapa_material})

    if busca_material_etapa(procura_lacre, "50001070", etapa):
        btn_busca_material(tb_materiais, session, "50001070")
        quantidade = tb_materiais.GetCellValue(
            tb_materiais.CurrentCellRow,
            "QUANT",
        )
        qtd_float = float(quantidade.replace(",", "."))
        if qtd_float >= max_qtd and not lacre_estoque.empty:
            tb_materiais.modifyCell(
                tb_materiais.CurrentCellRow,
                "QUANT",
                "1",
            )
            tb_materiais.setCurrentCell(
                tb_materiais.CurrentCellRow,
                "QUANT",
            )

        # Caso lacre sem estoque
        if lacre_estoque.empty:
            try:
                tb_materiais.modifyCell(
                    tb_materiais.CurrentCellRow,
                    "QUANT",
                    "0",
                )
                tb_materiais.setCurrentCell(
                    tb_materiais.CurrentCellRow,
                    "QUANT",
                )
                tb_materiais.modifyCheckBox(
                    tb_materiais.CurrentCellRow,
                    "ELIMINADO",
                    True,
                )
            except pywintypes.com_error:
                logger.exception("Erro ao modificar a célula no módulo lacre material.")

    if not busca_material_etapa(procura_lacre, "50001070", etapa) and not lacre_estoque.empty:
        tb_materiais.InsertRows(str(ultima_linha_material))
        tb_materiais.modifyCell(
            ultima_linha_material,
            "ETAPA",
            etapa,
        )
        tb_materiais.modifyCell(
            ultima_linha_material,
            "MATERIAL",
            "50001070",
        )
        tb_materiais.modifyCell(
            ultima_linha_material,
            "QUANT",
            "1",
        )
        tb_materiais.setCurrentCell(
            ultima_linha_material,
            "QUANT",
        )
        ultima_linha_material = ultima_linha_material + 1
