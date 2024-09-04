"""Módulo Família Ligação Água Unitário."""

from __future__ import annotations

import typing

from python.src.lista_reposicao import dict_reposicao
from python.src.unitarios.localizador import btn_localizador

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


class LigacaoAgua:
    """Ramo de Ligações (Ramal) de água."""

    MND: tuple[str, str, str, str] = ("TA", "EI", "TO", "PO")
    # Ordem da tupla: [0] -> preço s/ fornecimento, [1] -> c/ fornecimento,
    # [2] Reposições -> (Cimentado, Especial e Asfalto Frio)
    # Códigos de preço para posição de rede do serviço pai
    CODIGOS: typing.ClassVar[dict[str, tuple[str, str, tuple[str, str, str]]]] = {
        "LAG_PA": ("456451", "456461", ("456471", "456472", "451495")),
        "LAG_MND": ("456491", "456492", ("456493", "456494", "451495")),
        "TRA_NV_PA": ("456801", "456811", ("456821", "456822", "451885")),
        "TRA_NV_MND": ("456881", "456882", ("456883", "456884", "451885")),
        "PNG_PA": ("456371", "456381", ("456391", "456392", "451415")),
        "PNG_MND": ("456411", "456881", ("456413", "456414", "451415")),
        "SUBST_PA": ("456451", "456461", ("456471", "456472", "451495")),
        "SUBST_MND": ("456491", "456492", ("456493", "456494", "451495")),
        "TRA_PREV_PA": ("456841", "456851", ("456861", "456862", "451895")),
        "TRA_PREV_MND": ("456891", "456892", ("456893", "456894", "451895")),
    }

    def __init__(
        self,
        etapa: str,
        corte: str,
        relig: str,
        reposicao: str,
        num_tse_linhas: int,
        etapa_reposicao: str,
        identificador: list[str],
        posicao_rede: str,
        profundidade: str,
        session: CDispatch,
        preco: CDispatch,
    ) -> None:
        """Construtor comum para Ramal de água.

        Args:
        ----
            etapa (str): Etapa pai do serviço.
            corte (str): Supressão
            relig (str): Restabelecimento
            reposicao (str): Serviço complementar
            num_tse_linhas (int): Total linhas do Grid
            etapa_reposicao (str): Etapa complementar
            identificador (list[str]): TSE, Etapa, Identificador para almoxarifado.py
            posicao_rede (str): Posição da Rede
            profundidade (str): Profundidade da rede
            session (CDispatch): Sessão SAP
            preco (CDispatch): Grid preço do SAP

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
        self._ramal = False
        self.preco = preco

    def supressao_ferrule(self) -> None:
        """Pagar Adicional de Supressão no ferrule/tomada."""
        btn_localizador(self.preco, self.session, "456506")
        self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
        self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
        self.preco.pressEnter()

    def reposicoes(self, cod_reposicao: tuple) -> None:
        """Reposições dos serviços de Ligação de água."""
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

    def _posicao_pagar(self, preco_tse: str) -> None:
        """Paga de acordo com a posição da rede."""
        if not self._ramal:
            btn_localizador(self.preco, self.session, preco_tse)
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()
            self._ramal = True

    def _repor(self, codigos_reposicao: tuple) -> None:
        if self.reposicao:
            self.reposicoes(codigos_reposicao)

    def _processar_operacao(self, tipo_operacao: str) -> None:
        codigo = (
            self.CODIGOS.get(tipo_operacao + "_PA") if self.posicao_rede == "PA" else self.CODIGOS.get(tipo_operacao + "_MND")
        )
        if codigo:
            self._posicao_pagar(codigo[0])
            if tipo_operacao == "SUBST":
                self.supressao_ferrule()
            self._repor(codigo[2])

    def ligacao_agua(self) -> None:
        """Ramal novo de água, avulsa."""
        if self.posicao_rede:
            self._processar_operacao("LAG")

    def tra_nv(self) -> None:
        """Troca de Ramal de água não visível."""
        if self.posicao_rede:
            self._processar_operacao("TRA_NV")

    def png(self) -> None:
        """Passado novo ramal para nova rede - Obra."""
        if self.posicao_rede:
            self._processar_operacao("PNG")

    def subst_agua(self) -> None:
        """Substituição de ramal de água, tem adicional de suprimir.

        o ferrule da rede.
        """
        if self.posicao_rede:
            self._processar_operacao("SUBST")

    def tra_prev(self) -> None:
        """Troca de ramal de água preventiva."""
        if self.posicao_rede:
            self._processar_operacao("TRA_PREV")
