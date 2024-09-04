"""Botão busca Aba Material."""

from __future__ import annotations

import typing

from pandas import DataFrame

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


def btn_busca_material(tb_materiais: CDispatch, session: CDispatch, codigo: str) -> None:
    """Botão localiza material."""
    tb_materiais.pressToolbarButton("&FIND")
    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = codigo
    session.findById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]").sendVKey(12)


def qtd_correta(tb_materiais: CDispatch, qtde: str) -> None:
    """Altera para quantidade correta informada."""
    tb_materiais.modifyCell(tb_materiais.CurrentCellRow, "QUANT", qtde)
    tb_materiais.setCurrentCell(tb_materiais.CurrentCellRow, "QUANT")


def qtd_max(material: str, estoque: DataFrame, limite: int, df_materiais: DataFrame) -> DataFrame:
    """Query no dataset da aba materiais."""
    if not estoque[estoque["Material"] == material].empty:
        return df_materiais[(df_materiais["Material"] == material) & (df_materiais["Quantidade"] > limite)]
    return DataFrame(columns=["Material", "Quantidade"]).astype({"Material": object, "Quantidade": int})
