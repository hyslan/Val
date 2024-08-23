"""Botão busca Aba Material."""


def btn_busca_material(tb_materiais, session, codigo) -> None:
    """Botão localiza material."""
    tb_materiais.pressToolbarButton("&FIND")
    session.findById(
        "wnd[1]/usr/txtGS_SEARCH-VALUE").text = codigo
    session.findById(
        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]").sendVKey(12)


def qtd_correta(tb_materiais, qtde) -> None:
    """Altera para quantidade correta informada."""
    tb_materiais.modifyCell(
        tb_materiais.CurrentCellRow, "QUANT", qtde)
    tb_materiais.setCurrentCell(
        tb_materiais.CurrentCellRow, "QUANT")


def qtd_max(material, estoque, limite, df_materiais):
    """Query no dataset da aba materiais."""
    if not estoque[estoque["Material"] == material].empty:
        return df_materiais[df_materiais["Material"] ==
                                 material].query(f"`Quantidade` > {limite}")
    return estoque[estoque["Material"] == material].empty
