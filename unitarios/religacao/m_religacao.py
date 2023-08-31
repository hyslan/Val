# religacao.py
'''Módulo Família Religação Unitário.'''
# pylint: disable=W0611
import sys
from lista_reposicao import dict_reposicao
from sap_connection import connect_to_sap
session = connect_to_sap()


class Religacao:
    '''Classe de Religação Unitário.'''
    @staticmethod
    def restabelecida(_,
                      relig,
                      reposicao,
                      num_tse_linhas,
                      etapa_reposicao,
                      posicao_rede,
                      profundidade,
                      ):
        '''Método para definir de qual forma foi restabelecida e 
        pagar de acordo com as informações dadas, caso contrário,
        pagar como ramal se tiver reposição ou cavalete.'''
        if relig == 'CAVALETE':
            print("Iniciando processo de pagar RELIG CV - Código: 456037")
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
                    if sap_preco == str(456037):
                        # Marca pagar na TSE
                        preco.modifyCell(n_preco, "QUANT", "1")
                        preco.setCurrentCell(n_preco, "QUANT")
                        preco.pressEnter()
                        print("Pago 1 UN de RELIG CV - CODIGO: 456037")
                        break

        elif relig == 'RAMAL PEAD' or reposicao:
            print(
                "Iniciando processo de pagar RELIG RAMAL AG S/REP - Código: 456039")
            preco = session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
            preco.GetCellValue(0, "NUMERO_EXT")

            ramal = False
            num_linhas_visiveis = 48
            num_precos_linhas = preco.RowCount
            n_preco = 0  # índice para itens de preço
            contador_pg = 0
            # Function lambda com list compreenhension para matriz de reposições.
            if reposicao:
                rep_com_etapa = [(x, y)
                                 for x, y in zip(reposicao, etapa_reposicao)]

                for pavimento in rep_com_etapa:
                    operacao_rep = pavimento[1]
                    if operacao_rep == '0':
                        operacao_rep = '0010'
                    # 0 é tse da reposição;
                    # 1 é etapa da tse da reposição;
                    if pavimento[0] in dict_reposicao['cimentado']:
                        preco_reposicao = str(456041)
                        txt_reposicao = (
                            "Pago 1 UN de LRP CIM RELIGACAO DE LIGACAO SUPR - CODIGO: 456041")
                    if pavimento[0] in dict_reposicao['especial']:
                        preco_reposicao = str(456042)
                        txt_reposicao = (
                            "Pago 1 UN de LRP ESP RELIGACAO DE LIGACAO SUPR - CODIGO: 456042")
                    if pavimento[0] in dict_reposicao['asfalto_frio']:
                        preco_reposicao = str(451043)
                        txt_reposicao = ("Pago 1 UN de LPB ASF SUPRE  LAG COMPX C"
                                         + " - CODIGO: 456042")
                for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
                    if contador_pg >= num_tse_linhas:
                        break
                    sap_preco = preco.GetCellValue(
                        n_preco, "NUMERO_EXT")
                    item_preco = preco.GetCellValue(
                        n_preco, "ITEM")
                    n_etapa = preco.GetCellValue(
                        n_preco, "ETAPA")

                    if ramal is False:
                        if sap_preco == str(456039):
                            # Marca pagar na TSE
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            print(
                                "Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                            contador_pg += 1
                            ramal = True

                        # 660 é módulo despesa.

                    if sap_preco == preco_reposicao and item_preco == '660' \
                            and n_etapa == operacao_rep:
                        preco.modifyCell(n_preco, "QUANT", "1")
                        preco.setCurrentCell(n_preco, "QUANT")
                        print(txt_reposicao)
                        contador_pg += 1

                    # 1820 é módulo despesa para cimentado e especial.
                    if preco_reposicao in ('456041', '456042'):
                        if sap_preco == preco_reposicao and item_preco == '1820' \
                                and n_etapa == operacao_rep:
                            preco.modifyCell(n_preco, "QUANT", "1")
                            preco.setCurrentCell(n_preco, "QUANT")
                            print(txt_reposicao)
                            contador_pg += 1

                    # Rola uma página para baixo para carregar mais rows.
                    if n_preco >= num_linhas_visiveis and num_precos_linhas > 96:
                        preco.currentCellRow = num_linhas_visiveis = 48 * 2
                    if n_preco > num_linhas_visiveis * 2 and num_precos_linhas > 192:
                        preco.currentCellRow = num_linhas_visiveis = 48 * 4

            if ramal is False:
                num_precos_linhas = preco.RowCount
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

                    if sap_preco == str(456039):
                        # Marca pagar na TSE
                        preco.modifyCell(n_preco, "QUANT", "1")
                        preco.setCurrentCell(n_preco, "QUANT")
                        print(
                            "Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                        contador_pg += 1
                        ramal = True

        elif relig == 'FERRULE':
            print(
                "Iniciando processo de pagar RELIG  TMD AG  S/REP - Código: 456040")
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
                    if sap_preco == str(456040):
                        # Marca pagar na TSE
                        preco.modifyCell(n_preco, "QUANT", "1")
                        preco.setCurrentCell(n_preco, "QUANT")
                        preco.pressEnter()
                        print(
                            "Pago 1 UN de RELIG  TMD AG  S/REP - CODIGO: 456040")
                        break
                    else:
                        print(
                            f"Código de preço: {sap_preco} , Linha: {n_preco} - Não Selecionado")
        else:
            print("Religação não informada. \n Pagando como RELIG CV.")
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
                    if sap_preco == str(456037):
                        # Marca pagar na TSE
                        preco.modifyCell(n_preco, "QUANT", "1")
                        preco.setCurrentCell(n_preco, "QUANT")
                        preco.pressEnter()
                        print("Pago 1 UN de RELIG CV - CODIGO: 456037")
                        break

        # Confirmação da precificação.
        preco.pressEnter()
