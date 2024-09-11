# religacao.py
"""Módulo Família Religação Unitário."""

from __future__ import annotations

import logging
import typing

import pywintypes

from python.src.lista_reposicao import dict_reposicao
from python.src.unitarios.localizador import btn_localizador

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch

logger = logging.getLogger(__name__)


class Religacao:
    """Classe de Religação Unitário."""

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
        """Construtor de Religação.

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

    def restabelecida(self) -> None:
        """Método para definir de qual forma foi restabelecida.

        E pagar de acordo com as informações dadas, caso contrário,
        pagar como ramal se tiver reposição ou cavalete.
        """
        try:
            if self.relig == "CAVALETE":
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456037")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    return

            if self.relig in ("RAMAL PEAD", "PASSEIO") or self.reposicao:
                self.preco.GetCellValue(0, "NUMERO_EXT")

                ramal = False
                contador_pg = 0
                # Function lambda com list compreenhension para matriz de reposições.
                if self.reposicao:
                    rep_com_etapa = list(zip(self.reposicao, self.etapa_reposicao, strict=False))

                    for pavimento in rep_com_etapa:
                        operacao_rep = pavimento[1]
                        if operacao_rep == "0":
                            operacao_rep = "0010"
                        # 0 é tse da reposição;
                        # 1 é etapa da tse da reposição;
                        if pavimento[0] in dict_reposicao["cimentado"]:
                            preco_reposicao = str(456041)
                        if pavimento[0] in dict_reposicao["especial"]:
                            preco_reposicao = str(456042)
                        if pavimento[0] in dict_reposicao["asfalto_frio"]:
                            preco_reposicao = str(451043)

                        if contador_pg >= self.num_tse_linhas:
                            return

                        if ramal is False:
                            btn_localizador(self.preco, self.session, "456039")
                            # Marca pagar na TSE
                            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                            self.preco.pressEnter()
                            contador_pg += 1
                            ramal = True

                            # 660 é módulo despesa.
                            btn_localizador(self.preco, self.session, preco_reposicao)
                            item_preco = self.preco.GetCellValue(
                                self.preco.CurrentCellRow,
                                "ITEM",
                            )
                            if item_preco in ("300", "1790"):
                                self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                                self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                                self.preco.pressEnter()
                                contador_pg += 1

                        # 1820 é módulo despesa para cimentado e especial.
                        if preco_reposicao in ("456041", "456042"):
                            btn_localizador(self.preco, self.session, preco_reposicao)
                            item_preco = self.preco.GetCellValue(
                                self.preco.CurrentCellRow,
                                "ITEM",
                            )
                            if item_preco in ("1820", "1830", "670", "5310"):
                                self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                                self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                                self.preco.pressEnter()
                                contador_pg += 1
                    return

                if ramal is False:
                    contador_pg = 0
                    btn_localizador(self.preco, self.session, "456039")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    contador_pg += 1
                    ramal = True
                    return

            if self.relig in ("FERRULE", "TOMADA/FERRULE"):
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456040")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    return

            if self.relig is None:
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456037")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    return

        except pywintypes.com_error:
            logger.exception("Erro no módulo Religação.")

        # Confirmação da precificação.
        self.preco.pressEnter()
