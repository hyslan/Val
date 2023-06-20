# cavalete.py
from sap_connection import connect_to_sap
session = connect_to_sap()


class Cavalete:
    @staticmethod
    # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
    def TrocaPeCvPrev(*args):
        print("Iniciando processo de pagar ADC  TRC PREV PE CV - Código: 456856")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0  # índice para itens de preço
            for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if sap_preco == str(456856):  # THD  ATE 10M3/H PREV
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de ADC  TRC PREV PE CV - CODIGO: 456856")
                    break
                else:
                    print(
                        f"Código de preço: {sap_preco} , Linha: {n_preco} - Não Selecionado")

    @staticmethod
    def TrocaCvKit(*args):  # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo de pagar TROCA DE CAVALETE (KIT)  - Código: 456011")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0  # índice para itens de preço
            for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if sap_preco == str(456011):
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de TCV SF - CODIGO: 456011")
                    break
                else:
                    print(
                        f"Código de preço: {sap_preco} , Linha: {n_preco} - Não Selecionado")
