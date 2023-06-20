# supressao.py
'''Módulo Família Supressão Unitário.'''
# Bibliotecas
# pylint: disable=W0611
import sys
from lista_reposicao import dict_reposicao
from sap_connection import connect_to_sap
session = connect_to_sap()


class Corte:
    '''Classe de Reposição Unitário.'''
    @staticmethod
    def supressao(corte, _, reposicao, num_tse_linhas, etapa_reposicao):
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
                    for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                        sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                        if sap_preco == str(456033):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de SUPR CV - CODIGO: 456033")
                            break

            elif corte == 'RAMAL' or reposicao:
                print(
                    "Iniciando processo de pagar SUPR  RAMAL AG  S/REP - Código: 456035")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if reposicao in dict_reposicao['cimentado']:
                    preco_reposicao = str(456041)
                    txt_reposicao = ("Pago 1 UN de LRP CIM RELIGACAO"
                                     + "DE LIGACAO SUPR - CODIGO: 456041")
                elif reposicao in dict_reposicao['especial']:
                    preco_reposicao = str(456042)
                    txt_reposicao = ("Pago 1 UN de LRP ESP RELIGACAO"
                                     + "DE LIGACAO SUPR - CODIGO: 456042")
                elif reposicao in dict_reposicao['asfalto_frio']:
                    preco_reposicao = str(451043)
                    txt_reposicao = ("Pago 1 UN de LPB ASF SUPRE  LAG COMPX C"
                                     + " - CODIGO: 456042")
                if preco is not None:
                    num_precos_linhas = preco.RowCount
                    print(
                        f"Qtd linhas em itens de preço: {num_precos_linhas}")
                    n_preco = 0  # índice para itens de preço
                    contador_pg = 0
                    for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                        if contador_pg >= num_tse_linhas:
                            break
                        sap_preco = preco.GetCellValue(
                            n_preco, "NUMERO_EXT")
                        item_preco = preco.GetCellValue(
                            n_preco, "ITEM")
                        n_etapa = preco.GetCellValue(
                            n_preco, "ETAPA")
                        if sap_preco == str(456035):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            print(
                                "Pago 1 UN de SUPR  RAMAL AG  S/REP - CODIGO: 456035")
                            contador_pg += 1
                        elif sap_preco == preco_reposicao and item_preco == '660' \
                                and preco_reposicao != str(451043):
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            print(txt_reposicao)
                            contador_pg += 1
                            # 660 é módulo despesa.
                        elif sap_preco == preco_reposicao and item_preco == '660' \
                                and n_etapa == etapa_reposicao:
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            print(txt_reposicao)
                            contador_pg += 1

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
                    for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                        sap_preco = preco.GetCellValue(
                            n_preco, "NUMERO_EXT")
                        if sap_preco == str(456034):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print(
                                "Pago 1 UN de SUPR  TMD AG  S/REP - CODIGO: 456034")
                            break
                        else:
                            print(
                                f"Código de preço: {sap_preco}, Linha: {n_preco} - Não Selecionado")
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
                    for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                        sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                        if sap_preco == str(456033):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de SUPR CV - CODIGO: 456033")
                            break
                        else:
                            print(
                                f"Código de preço: {sap_preco}, Linha: {n_preco} - Não Selecionado")
        # Confirmação da precificação.
        preco.pressEnter()
