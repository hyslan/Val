"""Módulo de busca código de preço."""

from __future__ import annotations

import typing

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


def btn_localizador(preco: CDispatch, session: CDispatch, codigo: str) -> None:
    """Objeto localizar."""
    preco.pressToolbarButton("&FIND")
    session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").Text = codigo
    session.findById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]").sendVKey(12)
