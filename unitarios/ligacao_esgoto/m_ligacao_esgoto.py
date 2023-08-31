'''Módulo Família Ligação Esgoto Unitário.'''
# pylint: disable=W0611
import sys
from lista_reposicao import dict_reposicao
from sap_connection import connect_to_sap
session = connect_to_sap()


class LigacaoEsgoto:
    '''Classe de Ligação Esgoto Unitário.'''
    @staticmethod
    def ligacao_esgoto(corte,
                       relig,
                       reposicao,
                       num_tse_linhas,
                       etapa_reposicao,
                       posicao_rede,
                       profundidade
                       ):
        '''Método para definir de qual forma foi a Ligação de esgoto e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a L.E .'''
        profundidade_float = float(profundidade.replace(",", "."))
        if profundidade_float <= 2.00:
            match posicao_rede:
                case 'PA':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: PA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456711)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456711")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456712)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456712")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451713)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TA LESG COMPX C"
                                                     + " - CODIGO: 451713")

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
                                    if sap_preco == str(456671):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")
                                        # LESG sem fornecimento Código: 456651
                                        # LESG sem fornecimento Código: 456671
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M PA AVUL S/REP - CODIGO: 456671")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456711', '456712'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456671):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    # LESG sem fornecimento Código: 456651
                                    # LESG sem fornecimento Código: 456671
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M PA AVUL S/REP - CODIGO: 456671")
                                    contador_pg += 1
                                    ramal = True
                case 'TA':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: TA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456711)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456711")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456712)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456712")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451713)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TA LESG COMPX C"
                                                     + " - CODIGO: 451713")

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
                                    if sap_preco == str(456672):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456652
                                        # LESG sem fornecimento Código: 456672
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M TA AVUL S/REP - CODIGO: 456672")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456711', '456712'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456672):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    # LESG sem fornecimento Código: 456652
                                    # LESG sem fornecimento Código: 456672
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M TA AVUL S/REP - CODIGO: 456672")
                                    contador_pg += 1
                                    ramal = True
                case 'EI':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: EI")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
                        num_linhas_visiveis = 48
                        num_precos_linhas = preco.RowCount
                        print(num_precos_linhas)
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
                                    preco_reposicao = str(456711)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456711")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456712)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456712")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451716)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M EX LESG COMPX C"
                                                     + " - CODIGO: 451716")

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
                                        if sap_preco == str(456673):
                                            # Marca pagar na TSE
                                            preco.modifyCell(
                                                n_preco, "QUANT", "1")
                                            preco.setCurrentCell(
                                                n_preco, "QUANT")
                                            # LESG sem fornecimento Código: 456653
                                            # LESG sem fornecimento Código: 456673
                                            print(
                                                "Pago 1 UN de LESG CER/PVC"
                                                + "ATE 2M EI AVUL S/REP - CODIGO: 456673")
                                            contador_pg += 1
                                            ramal = True
                                            preco.currentCellRow = num_linhas_visiveis = 64

                                    # 8140 é módulo Investimento.
                                    # print(preco_reposicao)
                                    # print(operacao_rep)

                                    if sap_preco == preco_reposicao \
                                            and n_etapa == operacao_rep:
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")
                                        print(txt_reposicao)
                                        contador_pg += 1

                                    # 5140 é módulo despesa para cimentado e especial.
                                    if preco_reposicao in ('456711', '456712'):
                                        if sap_preco == preco_reposicao \
                                                and n_etapa == operacao_rep:
                                            preco.modifyCell(
                                                n_preco, "QUANT", "1")
                                            preco.setCurrentCell(
                                                n_preco, "QUANT")
                                            print(txt_reposicao)
                                            contador_pg += 1

                                    # Rola uma página para baixo para carregar mais rows.
                                    if n_preco >= num_linhas_visiveis and num_precos_linhas > 48:
                                        num_linhas_visiveis += 1
                                        preco.currentCellRow = num_linhas_visiveis
                                    if n_preco > num_linhas_visiveis * 2 and num_precos_linhas > 192:
                                        num_linhas_visiveis += 1
                                        preco.currentCellRow = num_linhas_visiveis
                                    if num_linhas_visiveis >= num_precos_linhas - 2:
                                        break

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

                                    if sap_preco == str(456673):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")
                                        # LESG sem fornecimento Código: 456653
                                        # LESG sem fornecimento Código: 456673
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M EI AVUL S/REP - CODIGO: 456673")
                                        contador_pg += 1
                                        ramal = True
                case 'TO':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: TO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456711)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456711")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456712)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456712")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451719)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TO LESG COMPX C"
                                                     + " - CODIGO: 451719")

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
                                    if sap_preco == str(456674):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")
                                        # LESG sem fornecimento Código: 456654
                                        # LESG sem fornecimento Código: 456674
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M TO AVUL S/REP - CODIGO: 456674")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456711', '456712'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456674):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456654
                                    # LESG sem fornecimento Código: 456674
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M TO AVUL S/REP - CODIGO: 456674")
                                    contador_pg += 1
                                    ramal = True
                case 'PO':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: PO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456711)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456711")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456712)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456712")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451719)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TO LESG COMPX C"
                                                     + " - CODIGO: 451719")

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
                                    if sap_preco == str(456675):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456655
                                        # LESG sem fornecimento Código: 456675
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M PO AVUL S/REP - CODIGO: 456675")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456711', '456712'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456671):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456655
                                    # LESG sem fornecimento Código: 456675
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M PO AVUL S/REP - CODIGO: 456675")
                                    contador_pg += 1
                                    ramal = True
                case _:
                    return
        else:
            # Profundidade maior do que 2M.
            match posicao_rede:
                case 'PA':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: PA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456722)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456722")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456723")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451724)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF + DE 2M TA LESG COMPX C"
                                                     + " - CODIGO: 451724")

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
                                    if sap_preco == str(456676):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456656
                                        # LESG sem fornecimento Código: 456671
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M PA AVUL S/REP - CODIGO: 456676")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456722', '456723'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456671):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456651
                                    # LESG sem fornecimento Código: 456671
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M PA AVUL S/REP - CODIGO: 456671")
                                    contador_pg += 1
                                    ramal = True
                case 'TA':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: TA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456722)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456722")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456723")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451724)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF + DE 2M TA LESG COMPX C"
                                                     + " - CODIGO: 451724")

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
                                    if sap_preco == str(456677):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456657
                                        # LESG sem fornecimento Código: 456677
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M TA AVUL S/REP - CODIGO: 456677")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456722', '456723'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456677):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456657
                                    # LESG sem fornecimento Código: 456677
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M TA AVUL S/REP - CODIGO: 456677")
                                    contador_pg += 1
                                    ramal = True
                case 'EI':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: EI")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456722)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456722")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456723")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451727)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF + DE 2M EX LESG COMPX C"
                                                     + " - CODIGO: 451727")

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
                                    if sap_preco == str(456678):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456658
                                        # LESG sem fornecimento Código: 456678
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M EI AVUL S/REP - CODIGO: 456678")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.
                                if preco_reposicao in ('456722', '456723'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456678):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456658
                                    # LESG sem fornecimento Código: 456678
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M EI AVUL S/REP - CODIGO: 456678")
                                    contador_pg += 1
                                    ramal = True
                case 'TO':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: TO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456722)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456722")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456723")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451730)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF + DE 2M EX LESG COMPX C"
                                                     + " - CODIGO: 451730")

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
                                    if sap_preco == str(456679):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456659
                                        # LESG sem fornecimento Código: 456679
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M TO AVUL S/REP - CODIGO: 456679")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo para cimentado e especial.
                                if preco_reposicao in ('456722', '456723'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456679):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456659
                                    # LESG sem fornecimento Código: 456679
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M TO AVUL S/REP - CODIGO: 456679")
                                    contador_pg += 1
                                    ramal = True
                case 'PO':
                    print(
                        "Iniciando processo de pagar Ligação de Esgoto Posicão: PO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
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
                                    preco_reposicao = str(456722)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 456722")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 456723")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(451730)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF + 2M TO LESG COMPX C"
                                                     + " - CODIGO: 451730")

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
                                    if sap_preco == str(456680):
                                        # Marca pagar na TSE
                                        preco.modifyCell(n_preco, "QUANT", "1")
                                        preco.setCurrentCell(n_preco, "QUANT")

                                        # LESG sem fornecimento Código: 456660
                                        # LESG sem fornecimento Código: 456680
                                        print(
                                            "Pago 1 UN de LESG CER/PVC"
                                            + "ATE 2M PO AVUL S/REP - CODIGO: 456680")
                                        contador_pg += 1
                                        ramal = True

                                # 8140 é módulo Investimento.

                                if sap_preco == preco_reposicao and item_preco == '8140' \
                                        and n_etapa == operacao_rep:
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")
                                    print(txt_reposicao)
                                    contador_pg += 1

                                # 5140 é módulo para cimentado e especial.
                                if preco_reposicao in ('456722', '456723'):
                                    if sap_preco == preco_reposicao and item_preco == '5140' \
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

                                if sap_preco == str(456671):
                                    # Marca pagar na TSE
                                    preco.modifyCell(n_preco, "QUANT", "1")
                                    preco.setCurrentCell(n_preco, "QUANT")

                                    # LESG sem fornecimento Código: 456655
                                    # LESG sem fornecimento Código: 456675
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M PO AVUL S/REP - CODIGO: 456675")
                                    contador_pg += 1
                                    ramal = True
                case _:
                    return
        if posicao_rede or profundidade is None:
            return
        # Confirmação da precificação.
        preco.pressEnter()
