# hidrometro.py
'''Módulo Família Hidrômetro Unitário.'''
from sap_connection import connect_to_sap
session = connect_to_sap()


class Hidrometro:
    '''Classe Unitária de Hidrômetro.'''

    @staticmethod
    def troca_de_hidro_preventiva_agendada(*args):
        '''Troca de Hidro Preventiva - Código 456902'''
        print(
            "Iniciando processo de pagar TROCA DE HIDROMETRO PREVENTIVA - Código: 456902")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0  # índice para itens de preço
            for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if sap_preco == str(456902):  # THD  ATE 10M3/H PREV
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de THD  ATE 10M3/H PREV - CODIGO: 456902")
                    break

    @staticmethod
    def desinclinado_hidrometro(*args):
        '''Desinclinado Hidrômetro - Código 456022'''
        print("Iniciando processo de pagar COLOCADO HIDROMETRO NA POSIÇÃO CORRETA - Código: 456022")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0  # índice para itens de preço
            for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if sap_preco == str(456022):
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de DESINCL HD - CODIGO: 456022")
                    break

    @staticmethod
    def troca_de_hidro_corretivo(*args):
        '''Troca de Hidrômetro Corretivo - Código 456901'''
        print("Iniciando processo de pagar TROCA DE HIDROMETRO CORRETIVO - Código: 456901")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0  # índice para itens de preço
            for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if sap_preco == str(456901):
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de THD  ATE 10M3/H CORR - CODIGO: 456902")
                    break
