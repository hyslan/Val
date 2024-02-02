# hidrometro.py
'''Módulo Família Hidrômetro Unitário.'''
from src.unitarios.base import BaseUnitario
from src.unitarios.localizador import btn_localizador


class Hidrometro(BaseUnitario):
    '''Classe Unitária de Hidrômetro.'''

    def troca_de_hidro_preventiva_agendada(self):
        '''Troca de Hidro Preventiva - Código 456902'''
        print(
            "Iniciando processo de pagar TROCA DE HIDROMETRO PREVENTIVA - Código: 456902")
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, self.session, "456902")
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de THD PREV - CODIGO: 456902")

    def desinclinado_hidrometro(self):
        '''Desinclinado Hidrômetro - Código 456022'''
        print("Iniciando processo de pagar COLOCADO HIDROMETRO NA POSIÇÃO CORRETA - Código: 456022")
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, self.session, "456022")
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de DESINCL HD - CODIGO: 456022")

    def troca_de_hidro_corretivo(self):
        '''Troca de Hidrômetro Corretivo - Código 456901'''
        print("Iniciando processo de pagar TROCA DE HIDROMETRO CORRETIVO - Código: 456901")
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, self.session, "456901")
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de THD  ATE 10M3/H CORR - CODIGO: 456901")
