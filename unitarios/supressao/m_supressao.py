# supressao.py
'''Módulo Família Supressão Unitário.'''
# Bibliotecas
import sys
import lista_reposicao
from sap_connection import connect_to_sap
session = connect_to_sap()


class Corte:
    '''Classe de Reposição Unitário.'''
    @staticmethod
    def Supressao(corte, _, reposicao):
        '''Método para definir de qual forma foi suprimida e 
        pagar de acordo com as informações dadas, caso contrário,
        pagar como ramal se tiver reposição ou cavalete.'''
        if corte is not None:
            if corte == 'CAVALETE':
                print("Iniciando processo de pagar SUPR CV - Código: 456033")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    num_precos_linhas = preco.RowCount
                    print(
                        f"Qtd linhas em itens de preço: {num_precos_linhas}")
                    n_preco = 0  # índice para itens de preço
                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                        SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                        if SAP_preco == str(456033):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de SUPR CV - CODIGO: 456033")
                            break

            elif corte == 'RAMAL' or reposicao is not None:
                print(
                    "Iniciando processo de pagar SUPR  RAMAL AG  S/REP - Código: 456035")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    num_precos_linhas = preco.RowCount
                    print(
                        f"Qtd linhas em itens de preço: {num_precos_linhas}")
                    n_preco = 0  # índice para itens de preço
                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                        SAP_preco = preco.GetCellValue(
                            n_preco, "NUMERO_EXT")
                        if SAP_preco == str(456035):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print(
                                "Pago 1 UN de SUPR  RAMAL AG  S/REP - CODIGO: 456035")
                            if reposicao in lista_reposicao.cimentado and SAP_preco == str(456041):
                                preco.modifyCell(n_preco, "QUANT", "1")
                                preco.setCurrentCell(n_preco, "QUANT")
                                preco.pressEnter()
                                print(
                                    "Pago 1 UN de LRP CIM RELIGACAO DE LIGACAO SUPR - CODIGO: 456041")
                                break
                            elif reposicao in lista_reposicao.especial and SAP_preco == str(456042):
                                preco.modifyCell(n_preco, "QUANT", "1")
                                preco.setCurrentCell(n_preco, "QUANT")
                                preco.pressEnter()
                                print(
                                    "Pago 1 UN de LRP ESP RELIGACAO DE LIGACAO SUPR - CODIGO: 456042")
                                break
                            elif reposicao in lista_reposicao.asfalto_frio:
                                print("Tem asfalto frio na bagaça!")
                                sys.exit()
                            else:
                                print(
                                    f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")

            elif corte == 'FERRULE':
                print(
                    "Iniciando processo de pagar SUPR  TMD AG  S/REP - Código: 456034")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    num_precos_linhas = preco.RowCount
                    print(
                        f"Qtd linhas em itens de preço: {num_precos_linhas}")
                    n_preco = 0  # índice para itens de preço
                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                        SAP_preco = preco.GetCellValue(
                            n_preco, "NUMERO_EXT")
                        if SAP_preco == str(456034):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print(
                                "Pago 1 UN de SUPR  TMD AG  S/REP - CODIGO: 456034")
                            break
                        else:
                            print(
                                f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")
            else:
                print("Corte não informado. \nPagando como SUPR CV.")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    num_precos_linhas = preco.RowCount
                    print(
                        f"Qtd linhas em itens de preço: {num_precos_linhas}")
                    n_preco = 0  # índice para itens de preço
                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                        SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                        if SAP_preco == str(456033):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de SUPR CV - CODIGO: 456033")
                            break
                        else:
                            print(
                                f"Código de preço: {SAP_preco}, Linha: {n_preco} - Não Selecionado")
