# base.py
"""Base construtora de classes."""

from __future__ import annotations

import typing

from python.src.unitarios.interfaces import UnitarioInterface

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


class BaseUnitario(UnitarioInterface):
    """Construtor comum."""

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
    ) -> None:
        """Construtor comum.

        Args:
        ----
            etapa (str): Etapa pai do serviço.
            corte (str): Supressão
            relig (str): Restabelecimento
            reposicao (list[str]): Serviço complementar
            num_tse_linhas (int): Total linhas do Grid
            etapa_reposicao (list[str]): Etapa complementar
            identificador (list[str]): TSE, Etapa, Identificador para almoxarifado.py
            posicao_rede (str): Posição da Rede
            profundidade (str): Profundidade da rede
            session (CDispatch): Sessão SAP
            preco (CDispatch): Grid preço do SAP

        """
        self.corte = corte
        self.relig = relig
        self.reposicao = reposicao
        self.num_tse_linhas = num_tse_linhas
        self.etapa_reposicao = etapa_reposicao
        self.posicao_rede = posicao_rede
        self.profundidade = profundidade
        self.session = session
        self.identificador = identificador
        self.etapa = etapa
        self.preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            "ZSBMM_VALORACAO_NAPI:9020/cntlCC_ITEM_PRECO/shellcont/shell",
        )

    def processar(self) -> None:
        """Processar."""

    def pagar(self) -> None:
        """Pagar."""

    def _processar_operacao(self, tipo_operacao: str) -> None:
        """Processar Código de preço."""
