"""Módulo de consulta estoque de materiais."""

from __future__ import annotations

import contextlib
import logging
import sys
import time
import typing
from pathlib import Path

import pandas as pd
import pywintypes
import xlwings as xw

from python.src import sap

logger = logging.getLogger(__name__)
if typing.TYPE_CHECKING:
    from pandas import DataFrame
    from win32com.client import CDispatch


def estoque(session: CDispatch, contrato: str, n_con: int) -> DataFrame:
    """Função para consultar estoque."""
    caminho = Path.cwd() / "sheets"
    caminho_str = str(caminho)
    try:
        session.StartTransaction("MBLB")
        frame = session.findById("wnd[0]")
        frame.findByid("wnd[0]/usr/ctxtLIFNR-LOW").text = contrato
        frame.SendVkey(8)
        frame.sendVKey(42)  # Lista Detalhada
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        grid.contextMenu()
        grid.selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = caminho_str
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"estoque_{contrato}.XLSX"
        session.findById("wnd[1]").sendVKey(11)  # Substituir
        time.sleep(2)
        materiais = pd.read_excel(
            caminho.joinpath(f"estoque_{contrato}.XLSX"),
            sheet_name="Sheet1",
            usecols=[
                "Material",
                "Texto breve material",
                "Utilização livre",
            ],
        )
        materiais = materiais.dropna()
        materiais["Material"] = materiais["Material"].astype(int).astype(str)
        sessao = sap
        con = sessao.connection_object(n_con)
        total_sessoes = sessao.contar_sessoes(n_con)
        max_sessoes = 6
        if total_sessoes != max_sessoes:
            with contextlib.suppress(Exception):
                con.CloseSession(session.ID)
            time.sleep(3)

        try:
            time.sleep(8)
            book = xw.Book(f"estoque_{contrato}.xlsx")
            book.app.quit()
        except xw.XlwingsError:
            pass
    except pywintypes.com_error:
        logger.exception("Erro ao consultar estoque.")
        sys.exit(1)

    return materiais
