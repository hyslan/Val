'''Módulo Família Poço Unitário.'''
from sap_connection import connect_to_sap
session = connect_to_sap()


class Poco:
    '''Classe Unitária de Poço.'''

    @staticmethod
    def troca_de_caixa_de_parada(*args):
        '''Troca de Caixa de Parada - Código 456112'''
        print(
            "Iniciando processo de pagar TROCA DE CAIXA DE PARADA - Código: 456112")
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
                # DETCX  DESC, NIV CX PARADA   S/REP
                if sap_preco == str(456112):
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print(
                        "Pago 1 UN de DETCX  DESC, NIV CX PARADA   S/REP - CODIGO: 456112")
                    break
