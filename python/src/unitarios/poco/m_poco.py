"""Módulo Família Poço Unitário."""

from __future__ import annotations

import typing

from python.src.lista_reposicao import dict_reposicao
from python.src.unitarios.localizador import btn_localizador

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


class Poco:
    """Classe Unitária de Poço."""

    CODIGOS: typing.ClassVar[dict[str, tuple[str, str, tuple[str, str, str]]]] = {
        "NIV_CX_PARADA": ("456111", ("050401", "050402", "451123")),
        "TROCA_CX_PARADA": ("456112", ("050401", "050402", "451123")),
        "NIVELAMENTO": ("456206", "451208", ("456207", "456207", "451208")),
    }

    def __init__(
        self,
        etapa: str,
        corte: str,
        relig: str,
        reposicao: str,
        num_tse_linhas: int,
        etapa_reposicao: str,
        identificador: str,
        posicao_rede: str,
        profundidade: str,
        session: CDispatch,
        preco: CDispatch,
    ) -> None:
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

    def reposicoes(self, cod_reposicao: tuple) -> None:
        """Reposições dos serviços de Poço."""
        rep_com_etapa = list(zip(self.reposicao, self.etapa_reposicao, strict=False))

        for pavimento in rep_com_etapa:
            operacao_rep = pavimento[1]
            if operacao_rep == "0":
                operacao_rep = "0010"
            # 0 é tse da reposição;
            # 1 é etapa da tse da reposição;
            if pavimento[0] in dict_reposicao["cimentado"]:
                preco_reposicao = cod_reposicao[0]
            if pavimento[0] in dict_reposicao["especial"]:
                preco_reposicao = cod_reposicao[1]
            if pavimento[0] in dict_reposicao["asfalto_frio"]:
                preco_reposicao = cod_reposicao[2]

            # 4220 é módulo Investimento.

            btn_localizador(self.preco, self.session, preco_reposicao)
            n_etapa = self.preco.GetCellValue(self.preco.CurrentCellRow, "ETAPA")

            if n_etapa != operacao_rep:
                self.preco.pressToolbarButton("&FIND")
                self.session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                self.session.findById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                self.session.findById("wnd[1]").sendVKey(0)
                self.session.findById("wnd[1]").sendVKey(0)
                self.session.findById("wnd[1]").sendVKey(12)

            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()

    def _repor(self, codigos_reposicao: tuple) -> None:
        if self.reposicao:
            self.reposicoes(codigos_reposicao)

    def _pagar(self, preco_tse: str) -> None:
        """Pagar serviço."""
        btn_localizador(self.preco, self.session, preco_tse)
        self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
        self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
        self.preco.pressEnter()

    def _processar_operacao(self, tipo_operacao: str) -> None:
        codigo = self.CODIGOS.get(tipo_operacao)
        if codigo:
            if tipo_operacao == "NIVELAMENTO":
                self._pagar(codigo[1])
                return

            self._pagar(codigo[0])
            self._repor(codigo[1])

    def niv_cx_parada(self) -> None:
        """Método Nivelamento de Caixa de Parada."""
        self._processar_operacao("NIV_CX_PARADA")

    def troca_de_caixa_de_parada(self) -> None:
        """Troca/Descobrimento de Caixa de Parada, Válvula de Rede de Água - Código 456112."""
        self._processar_operacao("TROCA_CX_PARADA")

    def nivelamento(self) -> None:
        """Nivelamentos com e sem reposição."""
        self._processar_operacao("NIVELAMENTO")
