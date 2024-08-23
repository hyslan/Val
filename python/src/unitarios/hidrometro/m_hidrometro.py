# hidrometro.py
"""Módulo Família Hidrômetro Unitário."""
from python.src.unitarios.localizador import btn_localizador


class Hidrometro:
    """Classe Unitária de Hidrômetro."""

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
