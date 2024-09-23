# transact_zsbmm216.py
"""Módulo de Contrato."""

from __future__ import annotations

import threading
import typing

import pythoncom
import pywintypes
import win32com.client as win32
from rich.console import Console

from python.src import sap

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch
# Adicionando um Lock
lock = threading.Lock()
console = Console()


class Transacao:
    """Classe operadora da transação 216."""

    def __init__(self, contrato: str, municipio: str, session: CDispatch) -> None:
        """Construtor de Transacao.

        Args:
        ----
            contrato (str): Número do Contrato
            municipio (str): Código do Município
            session (CDispatch): Sessão do SAP

        """
        self._contrato = contrato
        self.session = session
        self._municipio = municipio

    @property
    def municipio(self) -> str:
        """Getter para municipio."""
        return self._municipio

    @municipio.setter
    def municipio(self, cod: str) -> None:
        """Setter for municipio."""
        if isinstance(cod, str):
            self._municipio = cod

    @property
    def contrato(self) -> str:
        """Getter for contrato.

        Returns
        -------
            str: Número do Contrato.

        """
        return self._contrato

    @contrato.setter
    def contrato(self, cod: str) -> None:
        if isinstance(cod, str):
            self._contrato = cod

    def run_transacao(self, ordem: str, tipo: str = "individual", n_con: int = 0) -> None:
        """Run thread ZSBMM216.

        E faz a transação a transação com o respectivo contrato.
        """

        def t_transacao(session_id: pywintypes.HANDLE) -> None:  # type: ignore
            """Transação preenchida ZSBMM216 - Contrato NOVASP."""
            nonlocal ordem

            # Seção Crítica - uso do Lock
            with lock:
                try:
                    pythoncom.CoInitialize()
                    gui = win32.Dispatch(
                        pythoncom.CoGetInterfaceAndReleaseStream(session_id, pythoncom.IID_IDispatch),
                    )
                    gui.StartTransaction("ZSBMM216")
                    if tipo == "consulta":
                        gui.findById("wnd[0]/usr/radRB_CON").Select()

                    # Unidade Administrativa
                    gui.findById("wnd[0]/usr/ctxtP_UND").Text = "1093"
                    # Contrato
                    gui.findById("wnd[0]/usr/ctxtP_CONT").Text = self._contrato
                    gui.findById("wnd[0]/usr/ctxtP_MUNI").Text = self._municipio  # Cidade
                    sap_ordem = gui.findById("wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
                    sap_ordem.Text = ordem
                    gui.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

                except pywintypes.com_error as erro_216:
                    console.print(f"Erro na ZSBMM216: {erro_216}")
                    console.print_exception(show_locals=True)

        try:
            pythoncom.CoInitialize()
            session_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, self.session)  # type: ignore
            # Start
            thread = threading.Thread(target=t_transacao, kwargs={"session_id": session_id})
            thread.start()
            # Aguarde a thread concluir
            thread.join(timeout=300)
            if thread.is_alive():
                sap.fechar_conexao(n_con)

        except pywintypes.com_error as erro_thread216:
            console.print(f"Erro na thread da ZSBMM216: {erro_thread216}")
            console.print_exception(show_locals=True)
