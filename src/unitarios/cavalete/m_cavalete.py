'''Módulo Unitário de cavalete'''
# cavalete.py
from src.unitarios.localizador import btn_localizador
from src.unitarios.base import BaseUnitario


class Cavalete(BaseUnitario):
    '''Classe Cavalete unitário'''

    def instalado_lacre(self):
        '''Método Instalado Lacre Diversos'''
        print("Iniciando processo de pagar LACRHD - Código: 456021")
        codigo = "456021"
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, self.session, codigo)
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de LACRHD - CODIGO: 456021")

    def troca_pe_cv_prev(self):
        '''Método Troca de Pé de Cavalete Preventivo'''
        print("Iniciando processo de pagar ADC  TRC PREV PE CV - Código: 456856")
        codigo = "456856"
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, self.session, codigo)
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de ADC  TRC PREV PE CV - CODIGO: 456856")

    def troca_cv_kit(self):
        '''Método Troca de cavalete KIT'''
        print("Iniciando processo de pagar TROCA DE CAVALETE (KIT)  - Código: 456011")
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, self.session, "456011")
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de TCV SF - CODIGO: 456011")

    def troca_cv_por_uma(self):
        '''Método Troca de Cavalete Preventivo'''
        print("Iniciando processo de pagar SUBST CV POR UMA")
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            if self.reposicao:
                codigo = "457229"
                btn_localizador(preco, self.session, codigo)
                item_preco = preco.GetCellValue(
                    preco.CurrentCellRow, "ITEM"
                )
                if item_preco == '300':
                    preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    print(
                        "Pago 1 UN de SUBST CV POR UMA C/REP PASS - CODIGO: 457229")

            else:
                codigo = "457230"
                btn_localizador(preco, self.session, codigo)
                item_preco = preco.GetCellValue(
                    preco.CurrentCellRow, "ITEM"
                )
                if item_preco == '300':
                    preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    print(
                        "Pago 1 UN de SUBST CV POR UMA S/REP - CODIGO: 457230")
