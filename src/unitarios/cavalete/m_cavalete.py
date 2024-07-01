"""Módulo Unitário de cavalete"""
# cavalete.py
from src.unitarios.localizador import btn_localizador


class Cavalete:
    """Classe Cavalete unitário"""

    def __init__(self, etapa, corte, relig, reposicao, num_tse_linhas,
                 etapa_reposicao, identificador, posicao_rede,
                 profundidade, session, preco):
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

    def instalado_lacre(self):
        """Método Instalado Lacre Diversos"""
        print("Iniciando processo de pagar LACRHD - Código: 456021")
        codigo = "456021"
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, codigo)
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()
            print("Pago 1 UN de LACRHD - CODIGO: 456021")

    def troca_pe_cv_prev(self):
        """Método Troca de Pé de Cavalete Preventivo"""
        print("Iniciando processo de pagar ADC  TRC PREV PE CV - Código: 456856")
        codigo = "456856"
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, codigo)
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()
            print("Pago 1 UN de ADC  TRC PREV PE CV - CODIGO: 456856")

    def troca_cv_kit(self):
        """Método Troca de cavalete KIT"""
        print("Iniciando processo de pagar TROCA DE CAVALETE (KIT)  - Código: 456011")
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            btn_localizador(self.preco, self.session, "456011")
            self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()
            print("Pago 1 UN de TCV SF - CODIGO: 456011")

    def troca_cv_por_uma(self):
        """Método Troca de Cavalete Preventivo"""
        print("Iniciando processo de pagar SUBST CV POR UMA")
        self.preco.GetCellValue(0, "NUMERO_EXT")
        if self.preco is not None:
            if self.reposicao:
                codigo = "457229"
                btn_localizador(self.preco, self.session, codigo)
                item_preco = self.preco.GetCellValue(
                    self.preco.CurrentCellRow, "ITEM"
                )
                if item_preco in ('300', '310'):
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    print(
                        "Pago 1 UN de SUBST CV POR UMA C/REP PASS - CODIGO: 457229")

            else:
                codigo = "457230"
                btn_localizador(self.preco, self.session, codigo)
                item_preco = self.preco.GetCellValue(
                    self.preco.CurrentCellRow, "ITEM"
                )
                if item_preco in ('300', '310', '730'):
                    self.preco.modifyCell(self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    print(
                        "Pago 1 UN de SUBST CV POR UMA S/REP - CODIGO: 457230")
