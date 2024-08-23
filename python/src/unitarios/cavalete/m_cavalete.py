"""Módulo Unitário de cavalete."""
# cavalete.py
from python.src.unitarios.localizador import btn_localizador


class Cavalete:
    """Classe Cavalete unitário."""

    def __init__(self, etapa, corte, relig, reposicao, num_tse_linhas,
                 etapa_reposicao, identificador, posicao_rede,
                 profundidade, session, preco) -> None:
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

    def instalado_lacre(self) -> None:
        """Método Instalado Lacre Diversos."""
        codigo = "456021"
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, codigo)
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()

    def troca_pe_cv_prev(self) -> None:
        """Método Troca de Pé de Cavalete Preventivo."""
        codigo = "456856"
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, codigo)
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()

    def troca_cv_kit(self) -> None:
        """Método Troca de cavalete KIT."""
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, "456011")
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()

    def troca_cv_por_uma(self) -> None:
        """Método Troca de Cavalete Preventivo."""
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            if self.reposicao:
                codigo = "457229"
                btn_localizador(self.preco, self.session, codigo)
                item_preco = self.preco.GetCellValue(
                    self.preco.CurrentCellRow, "ITEM",
                )
                if item_preco in ("300", "310", "730", "1470", "4950",
                                  ):
                    self.preco.modifyCell(
                        self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(
                        self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()

            else:
                codigo = "457230"
                btn_localizador(self.preco, self.session, codigo)
                item_preco = self.preco.GetCellValue(
                    self.preco.CurrentCellRow, "ITEM",
                )
                if item_preco in ("300", "310", "730", "1470", "4950",
                                  ):
                    self.preco.modifyCell(
                        self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(
                        self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
