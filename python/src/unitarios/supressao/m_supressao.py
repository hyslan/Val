# supressao.py
"""Módulo Família Supressão Unitário."""

from __future__ import annotations

import logging
import typing

import pywintypes

from python.src.lista_reposicao import dict_reposicao
from python.src.unitarios.localizador import btn_localizador

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch

logger = logging.getLogger(__name__)


class Corte:
    """Classe de Reposição Unitário."""

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

    def supressao(self) -> None:
        """Método para definir de qual forma foi suprimida.

        E pagar de acordo com as informações dadas, caso contrário,
        pagar como ramal se tiver reposição ou cavalete.
        """
        try:
            if self.corte == "CAVALETE":
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456033")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    return

            if self.corte in ("RAMAL PEAD", "PASSEIO") or self.reposicao:
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
                            btn_localizador(self.preco, self.session, "456032")
                            self.preco.modifyCell(
                                self.preco.CurrentCellRow,
                                "QUANT",
                                "1",
                            )
                            self.preco.setCurrentCell(
                                self.preco.CurrentCellRow,
                                "QUANT",
                            )
                            self.preco.pressEnter()
                            contador_pg += 1
                            ramal = True

                            # 660 é módulo despesa.
                        btn_localizador(self.preco, self.session, preco_reposicao)
                        item_preco = self.preco.GetCellValue(
                            self.preco.CurrentCellRow,
                            "ITEM",
                        )
                        if item_preco == "660":
                            self.preco.modifyCell(
                                self.preco.CurrentCellRow,
                                "QUANT",
                                "1",
                            )
                            self.preco.setCurrentCell(
                                self.preco.CurrentCellRow,
                                "QUANT",
                            )
                            self.preco.pressEnter()
                            contador_pg += 1

                        # 1820 é módulo despesa para cimentado e especial.
                        if preco_reposicao in ("456041", "456042"):
                            btn_localizador(self.preco, self.session, preco_reposicao)
                            item_preco = self.preco.GetCellValue(
                                self.preco.CurrentCellRow,
                                "ITEM",
                            )
                            if item_preco in ("1820", "1830", "5310", "670"):
                                self.preco.modifyCell(
                                    self.preco.CurrentCellRow,
                                    "QUANT",
                                    "1",
                                )
                                self.preco.setCurrentCell(
                                    self.preco.CurrentCellRow,
                                    "QUANT",
                                )
                                self.preco.pressEnter()
                                contador_pg += 1

                if ramal is False:
                    contador_pg = 0
                    btn_localizador(self.preco, self.session, "456032")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    contador_pg += 1
                    ramal = True

                return

            if self.corte in ("FERRULE", "TOMADA/FERRULE"):
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456031")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    return

            if self.corte is None:
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456033")
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    return

        except pywintypes.com_error:
            logger.exception("Erro no módulo de Supressão.")
        # Confirmação da precificação.
        self.preco.pressEnter()
