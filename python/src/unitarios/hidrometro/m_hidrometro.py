# hidrometro.py
"""Módulo Família Hidrômetro Unitário."""

from __future__ import annotations

import typing

from python.src.unitarios.localizador import btn_localizador

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


class Hidrometro:
    """Classe Unitária de Hidrômetro."""

    def __init__(
        self,
        etapa: str,
        corte: str,
        relig: str,
        reposicao: list[str],
        num_tse_linhas: int,
        etapa_reposicao: list[str],
        identificador: list[str],
        posicao_rede: str,
        profundidade: str,
        session: CDispatch,
        preco: CDispatch,
    ) -> None:
        """Construtor de Supressão.

        Args:
        ----
            etapa (str): Etapa Pai
            corte (str): Onde foi feita a supressão
            relig (str): Onde foi realizado a religação
            reposicao (str): Etapa complementar
            num_tse_linhas (int): Count
            etapa_reposicao (str): Etapa do serviço complementar
            identificador (list[str]): TSE, Etapa, id match case do almoxarifado.py
            posicao_rede (str): Posição da Rede
            profundidade (str): Profundidade
            session (CDispatch): Sessão do SAPGUI
            preco (CDispatch): GRID de preço do SAP

        """
        self.etapa = etapa
        self.corte = corte
        self.relig = relig
        self.reposicao = reposicao
        self.num_tse_linhas = num_tse_linhas
        self.etapa_reposicao = etapa_reposicao
        self.posicao_rede = posicao_rede
        self.profundidade = profundidade
        self.session = session
        self.identificador = identificador
        self.preco = preco

    def troca_de_hidro_preventiva_agendada(self) -> None:
        """Troca de Hidro Preventiva - Código 456902."""
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, "456902")
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()

    def desinclinado_hidrometro(self) -> None:
        """Desinclinado Hidrômetro - Código 456022."""
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, "456022")
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()

    def troca_de_hidro_corretivo(self) -> None:
        """Troca de Hidrômetro Corretivo - Código 456901."""
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, "456901")
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()
