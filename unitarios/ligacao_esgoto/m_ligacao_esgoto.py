'''Módulo Família Ligação Esgoto Unitário.'''
# pylint: disable=W0611
import sys
from lista_reposicao import dict_reposicao
from sap_connection import connect_to_sap


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
        session = connect_to_sap()
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

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456651"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # LESG sem fornecimento Código: 456651
                                    # LESG sem fornecimento Código: 456671
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M PA AVUL S/REP - CODIGO: 456651")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456651"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456651
                            # LESG com fornecimento Código: 456671
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "ATE 2M PA AVUL S/REP - CODIGO: 456651")
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

                                if contador_pg >= num_tse_linhas:
                                    return

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456652"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # LESG sem fornecimento Código: 456652
                                    # LESG com fornecimento Código: 456672
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M TA AVUL S/REP - CODIGO: 456652")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456652"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456652
                            # LESG com fornecimento Código: 456672
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "ATE 2M TA AVUL S/REP - CODIGO: 456652")
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

                                if contador_pg >= num_tse_linhas:
                                    return

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456653"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # LESG sem fornecimento Código: 456653
                                    # LESG com fornecimento Código: 456673
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M EI AVUL S/REP - CODIGO: 456653")
                                    contador_pg += 1
                                    ramal = True

                                    # # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456653"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456653
                            # LESG sem fornecimento Código: 456673
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "ATE 2M EI AVUL S/REP - CODIGO: 456653")
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

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456654"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 456654
                                # LESG com fornecimento Código: 456674
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "ATE 2M TO AVUL S/REP - CODIGO: 456654")
                                contador_pg += 1
                                ramal = True

                                # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456654"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456654
                            # LESG com fornecimento Código: 456674
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "ATE 2M TO AVUL S/REP - CODIGO: 456654")
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

                                if contador_pg >= num_tse_linhas:
                                    return

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456655"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # LESG sem fornecimento Código: 456655
                                    # LESG sem fornecimento Código: 456675
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "ATE 2M PO AVUL S/REP - CODIGO: 456655")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456655"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456655
                            # LESG sem fornecimento Código: 456675
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "ATE 2M PO AVUL S/REP - CODIGO: 456655")
                            contador_pg += 1
                            ramal = True
                case _:
                    print("Profundidade não informada.")
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

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456656"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 456656
                                # LESG com fornecimento Código: 456676
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "+ 2M PA AVUL S/REP - CODIGO: 456656")
                                contador_pg += 1
                                ramal = True

                                # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456656"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456656
                            # LESG com fornecimento Código: 456676
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "+ 2M PA AVUL S/REP - CODIGO: 456656")
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

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456657"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 456657
                                # LESG sem fornecimento Código: 456677
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "+ 2M PA AVUL S/REP - CODIGO: 456657")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456657"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456657
                            # LESG sem fornecimento Código: 456677
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "+ 2M TA AVUL S/REP - CODIGO: 456657")
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

                                if contador_pg >= num_tse_linhas:
                                    break

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456658"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # LESG sem fornecimento Código: 456658
                                    # LESG sem fornecimento Código: 456678
                                    print(
                                        "Pago 1 UN de LESG CER/PVC"
                                        + "+ 2M EI AVUL S/REP - CODIGO: 456658")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456658"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456658
                            # LESG sem fornecimento Código: 456678
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "+ 2M EI AVUL S/REP - CODIGO: 456658")
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

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456659"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 456659
                                # LESG com fornecimento Código: 456679
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "+ 2M TO AVUL S/REP - CODIGO: 456659")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456659"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456659
                            # LESG com fornecimento Código: 456679
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "+ 2M TO AVUL S/REP - CODIGO: 456659")
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

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456660"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 456660
                                # LESG com fornecimento Código: 456680
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "+ 2M PO AVUL S/REP - CODIGO: 456660")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456660"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456660
                            # LESG com fornecimento Código: 456680
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "+ 2M PO AVUL S/REP - CODIGO: 456660")
                            contador_pg += 1
                            ramal = True
                case _:
                    return
        if posicao_rede or profundidade is None:
            return

    @staticmethod
    def png_esgoto(corte,
                   relig,
                   reposicao,
                   num_tse_linhas,
                   etapa_reposicao,
                   posicao_rede,
                   profundidade
                   ):
        '''Método para definir de qual forma foi a png e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a  PNG.'''
        session = connect_to_sap()
        match posicao_rede:
            case 'PA':
                print(
                    "Iniciando processo de pagar PNG Posicão: PA")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    ramal = False
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
                                preco_reposicao = str(456631)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP CIM  - CODIGO: 456631")
                            if pavimento[0] in dict_reposicao['especial']:
                                preco_reposicao = str(456632)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP ESP  - CODIGO: 456632")
                            if pavimento[0] in dict_reposicao['asfalto_frio']:
                                preco_reposicao = str(451633)
                                txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TA LESG COMPX C"
                                                 + " - CODIGO: 451633")

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456601"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # PNG sem fornecimento Código: 456601
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "ATE 2M PA AVUL S/REP - CODIGO: 456601")
                                contador_pg += 1
                                ramal = True

                            # 7860 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                    if ramal is False:
                        # Botão localizar
                        preco.pressToolbarButton("&FIND")
                        session.findById(
                            "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456601"
                        session.findById(
                            "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        # PNG sem fornecimento Código: 456601
                        print(
                            "Pago 1 UN de LESG CER/PVC"
                            + "ATE 2M PA AVUL S/REP - CODIGO: 456602")
                        contador_pg += 1
                        ramal = True
            case 'TA':
                print(
                    "Iniciando processo de pagar PNG Posicão: TA")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    ramal = False
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
                                preco_reposicao = str(456631)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP CIM  - CODIGO: 456631")
                            if pavimento[0] in dict_reposicao['especial']:
                                preco_reposicao = str(456632)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP ESP  - CODIGO: 456632")
                            if pavimento[0] in dict_reposicao['asfalto_frio']:
                                preco_reposicao = str(451633)
                                txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TA LESG COMPX C"
                                                 + " - CODIGO: 451633")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456602"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # PNG sem fornecimento Código: 456602
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "ATE 2M TA AVUL S/REP - CODIGO: 456602")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                    if ramal is False:
                        # Botão localizar
                        preco.pressToolbarButton("&FIND")
                        session.findById(
                            "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456602"
                        session.findById(
                            "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        # PNG sem fornecimento Código: 456602
                        print(
                            "Pago 1 UN de LESG CER/PVC"
                            + "ATE 2M TA AVUL S/REP - CODIGO: 456602")
                        contador_pg += 1
                        ramal = True
            case 'EI':
                print(
                    "Iniciando processo de pagar PNG Posicão: EI")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    ramal = False
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
                                preco_reposicao = str(456631)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP CIM  - CODIGO: 456631")
                            if pavimento[0] in dict_reposicao['especial']:
                                preco_reposicao = str(456632)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP ESP  - CODIGO: 456632")
                            if pavimento[0] in dict_reposicao['asfalto_frio']:
                                preco_reposicao = str(451636)
                                txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M EX LESG COMPX C"
                                                 + " - CODIGO: 451636")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456603"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # PNG sem fornecimento Código: 456603
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "ATE 2M EI AVUL S/REP - CODIGO: 456603")
                                contador_pg += 1
                                ramal = True

                                # # 7860 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 7940 é módulo despesa para cimentado e especial.

                    if ramal is False:
                        # Botão localizar
                        preco.pressToolbarButton("&FIND")
                        session.findById(
                            "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456603"
                        session.findById(
                            "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        # PNG sem fornecimento Código: 456603
                        print(
                            "Pago 1 UN de LESG CER/PVC"
                            + "ATE 2M EI AVUL S/REP - CODIGO: 456603")
                        contador_pg += 1
                        ramal = True
            case 'TO':
                print(
                    "Iniciando processo de pagar PNG Posicão: TO")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    ramal = False
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
                                preco_reposicao = str(456631)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP CIM  - CODIGO: 456631")
                            if pavimento[0] in dict_reposicao['especial']:
                                preco_reposicao = str(456632)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP ESP  - CODIGO: 456632")
                            if pavimento[0] in dict_reposicao['asfalto_frio']:
                                preco_reposicao = str(451639)
                                txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TO LESG COMPX C"
                                                 + " - CODIGO: 451639")

                        if contador_pg >= num_tse_linhas:
                            return

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456604"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456604
                            print(
                                "Pago 1 UN de LESG CER/PVC"
                                + "ATE 2M TO AVUL S/REP - CODIGO: 456604")
                            contador_pg += 1
                            ramal = True

                            # 7860 é módulo Investimento.

                        preco.pressToolbarButton("&FIND")
                        session.findById(
                            "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                        session.findById(
                            "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        n_etapa = preco.GetCellValue(
                            preco.CurrentCellRow, "ETAPA")

                        if not n_etapa == operacao_rep:
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)

                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        print(txt_reposicao)
                        contador_pg += 1

                        # 7940 é módulo despesa para cimentado e especial.

                    if ramal is False:
                        # Botão localizar
                        preco.pressToolbarButton("&FIND")
                        session.findById(
                            "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456604"
                        session.findById(
                            "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        # LESG sem fornecimento Código: 456604
                        print(
                            "Pago 1 UN de LESG CER/PVC"
                            + "ATE 2M TO AVUL S/REP - CODIGO: 456604")
                        contador_pg += 1
                        ramal = True
            case 'PO':
                print(
                    "Iniciando processo de pagar PNG Esgoto Posicão: PO")
                preco = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    ramal = False
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
                                preco_reposicao = str(456631)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP CIM  - CODIGO: 456631")
                            if pavimento[0] in dict_reposicao['especial']:
                                preco_reposicao = str(456632)
                                txt_reposicao = (
                                    "Pago 1 UN de LRP ESP  - CODIGO: 456632")
                            if pavimento[0] in dict_reposicao['asfalto_frio']:
                                preco_reposicao = str(451639)
                                txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TO LESG COMPX C"
                                                 + " - CODIGO: 451639")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456605"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 456605
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + "ATE 2M PO AVUL S/REP - CODIGO: 456605")
                                contador_pg += 1
                                ramal = True

                            # 7940 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 8140 é módulo despesa para cimentado e especial.

                    if ramal is False:
                        # Botão localizar
                        preco.pressToolbarButton("&FIND")
                        session.findById(
                            "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456605"
                        session.findById(
                            "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        # LESG sem fornecimento Código: 456605
                        print(
                            "Pago 1 UN de LESG CER/PVC"
                            + "ATE 2M PO AVUL S/REP - CODIGO: 456605")
                        contador_pg += 1
                        ramal = True
            case _:
                print("Posicão de rede não informada.")
                return

    @staticmethod
    def tre(corte,
            relig,
            reposicao,
            num_tse_linhas,
            etapa_reposicao,
            posicao_rede,
            profundidade
            ):
        '''Método para definir de qual forma foi a TRE e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a TRE .'''
        session = connect_to_sap()
        profundidade_float = float(profundidade.replace(",", "."))
        if profundidade_float <= 2.00:
            match posicao_rede:
                case 'PA':
                    print(
                        "Iniciando processo de pagar TRE Posicão: PA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457101)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457101")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457104)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457104")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452107)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TA TRE COMPX C"
                                                     + " - CODIGO: 452107")

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457001"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # TRE sem fornecimento Código: 457001
                                    # TRE com fornecimento Código: 457021
                                    print(
                                        "Pago 1 UN de TRE CER/PVC"
                                        + "ATE 2M PA AVUL S/REP - CODIGO: 456651")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457001"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457001
                            # TRE com fornecimento Código: 457021
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + "ATE 2M PA AVUL S/REP - CODIGO: 457001")
                            contador_pg += 1
                            ramal = True
                case 'TA':
                    print(
                        "Iniciando processo de pagar TRE Posicão: TA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457101)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457101")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457104)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457104")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452107)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TA TRE COMPX C"
                                                     + " - CODIGO: 452107")

                                if contador_pg >= num_tse_linhas:
                                    return

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457002"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # TRE sem fornecimento Código: 457002
                                    # TRE com fornecimento Código: 457022
                                    print(
                                        "Pago 1 UN de TRE CER/PVC"
                                        + "ATE 2M TA AVUL S/REP - CODIGO: 457002")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457002"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456652
                            # LESG com fornecimento Código: 457022
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + "ATE 2M TA AVUL S/REP - CODIGO: 457002")
                            contador_pg += 1
                            ramal = True
                case 'EI':
                    print(
                        "Iniciando processo de pagar TRE Posicão: EI")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457101)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457101")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457104)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457104")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452116)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M EX TRE COMPX C"
                                                     + " - CODIGO: 452116")

                                if contador_pg >= num_tse_linhas:
                                    return

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457003"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # TRE sem fornecimento Código: 457003
                                    # TRE com fornecimento Código: 457023
                                    print(
                                        "Pago 1 UN de TRE CER/PVC"
                                        + "ATE 2M EI AVUL S/REP - CODIGO: 457003")
                                    contador_pg += 1
                                    ramal = True

                                    # # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457003"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457003
                            # TRE sem fornecimento Código: 457023
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + "ATE 2M EI AVUL S/REP - CODIGO: 457003")
                            contador_pg += 1
                            ramal = True
                case 'TO':
                    print(
                        "Iniciando processo de pagar TRE Posicão: TO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457101)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457101")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457104)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457104")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452125)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TO TRE COMPX C"
                                                     + " - CODIGO: 452125")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457004"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457004
                                # TRE com fornecimento Código: 457024
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + "ATE 2M TO AVUL S/REP - CODIGO: 457004")
                                contador_pg += 1
                                ramal = True

                                # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457004"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457004
                            # TRE com fornecimento Código: 457024
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + "ATE 2M TO AVUL S/REP - CODIGO: 457004")
                            contador_pg += 1
                            ramal = True
                case 'PO':
                    print(
                        "Iniciando processo de pagar TRE Posicão: PO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457101)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457101")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457104)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457104")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452125)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF ATE 2M TO TRE COMPX C"
                                                     + " - CODIGO: 452125")

                                if contador_pg >= num_tse_linhas:
                                    return

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457005"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # TRE sem fornecimento Código: 457005
                                    # TRE com fornecimento Código: 457025
                                    print(
                                        "Pago 1 UN de TRE CER/PVC"
                                        + "ATE 2M PO AVUL S/REP - CODIGO: 457005")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                                # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457005"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457005
                            # TRE sem fornecimento Código: 457025
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + "ATE 2M PO AVUL S/REP - CODIGO: 457005")
                            contador_pg += 1
                            ramal = True
                case _:
                    print("Profundidade não informada.")
                    return
        elif 2.00 < profundidade_float <= 3.00:
            # Profundidade de 2M a 3M.
            match posicao_rede:
                case 'PA':
                    print(
                        "Iniciando processo de pagar TRE Posicão: PA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457102)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457102")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457105)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457105")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452110)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF DE 2M A 3M TA TRE COMPX C"
                                                     + " - CODIGO: 452110")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457006"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457006
                                # TRE com fornecimento Código: 457026
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + " 2M A 3M PA AVUL S/REP - CODIGO: 457006")
                                contador_pg += 1
                                ramal = True

                                # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457006"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457006
                            # TRE com fornecimento Código: 457002
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + "2M A 3M PA AVUL S/REP - CODIGO: 457006")
                            contador_pg += 1
                            ramal = True
                case 'TA':
                    print(
                        "Iniciando processo de pagar TRE Posicão: TA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457102)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457102")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457105)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457105")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452110)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF DE 2M A 3M TA TRE COMPX C"
                                                     + " - CODIGO: 452110")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457007"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457007
                                # TRE sem fornecimento Código: 457027
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + "2M A 3M TA AVUL S/REP - CODIGO: 457007")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457007"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457007
                            # TRE sem fornecimento Código: 457027
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 2M A 3M TA AVUL S/REP - CODIGO: 457007")
                            contador_pg += 1
                            ramal = True
                case 'EI':
                    print(
                        "Iniciando processo de pagar TRE Posicão: EI")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457102)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457102")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457105)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457105")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452119)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF DE 2M A 3M EX TRE COMPX C"
                                                     + " - CODIGO: 452119")

                                if contador_pg >= num_tse_linhas:
                                    break

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457008"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # TRE sem fornecimento Código: 457008
                                    # TRE sem fornecimento Código: 457028
                                    print(
                                        "Pago 1 UN de TRE CER/PVC"
                                        + "+ 2M EI AVUL S/REP - CODIGO: 457008")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457008"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457008
                            # TRE sem fornecimento Código: 457028
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 2M A 3M EI AVUL S/REP - CODIGO: 457008")
                            contador_pg += 1
                            ramal = True
                case 'TO':
                    print(
                        "Iniciando processo de pagar TRE Posicão: TO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457102)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457102")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457105)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457105")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452128)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF DE 2M A 3M EX LESG COMPX C"
                                                     + " - CODIGO: 452128")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457009"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457009
                                # TRE com fornecimento Código: 457029
                                print(
                                    "Pago 1 UN de LESG CER/PVC"
                                    + " 2M A 3M TO AVUL S/REP - CODIGO: 457009")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457009"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457009
                            # TRE com fornecimento Código: 457029
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 2M A 3M TO AVUL S/REP - CODIGO: 457009")
                            contador_pg += 1
                            ramal = True
                case 'PO':
                    print(
                        "Iniciando processo de pagar TRBM Posicão: PO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457102)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457102")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(457105)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457105")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452125)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF + 2M TO TRE COMPX C"
                                                     + " - CODIGO: 452125")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457010"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457010
                                # TRE com fornecimento Código: 457010
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + " 2M A 3M PO AVUL S/REP - CODIGO: 457010")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457010"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457010
                            # TRE com fornecimento Código: 457030
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 2M A 3M PO AVUL S/REP - CODIGO: 457010")
                            contador_pg += 1
                            ramal = True
                case _:
                    return
        elif 3.00 < profundidade_float <= 4.00:
            # Profundidade de 2M a 3M.
            match posicao_rede:
                case 'PA':
                    print(
                        "Iniciando processo de pagar TRE Posicão: PA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457103)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457103")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456457106723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457106")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452113)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF 3M A 4M TA TRE COMPX C"
                                                     + " - CODIGO: 452113")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457011"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # LESG sem fornecimento Código: 457011
                                # LESG com fornecimento Código: 457031
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + " 3M A 4M PA AVUL S/REP - CODIGO: 457011")
                                contador_pg += 1
                                ramal = True

                                # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457011"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 457011
                            # LESG com fornecimento Código: 457031
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 3M A 4M PA AVUL S/REP - CODIGO: 457011")
                            contador_pg += 1
                            ramal = True
                case 'TA':
                    print(
                        "Iniciando processo de pagar TRE Posicão: TA")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457103)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457103")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456457106723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457106")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452113)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF 3M A 4M TA TRE COMPX C"
                                                     + " - CODIGO: 452113")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457012"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457012
                                # TRE com fornecimento Código: 457031
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + " 3M A 4M TA AVUL S/REP - CODIGO: 457012")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457012"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457012
                            # TRE com fornecimento Código: 457032
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 3M A 4M TA AVUL S/REP - CODIGO: 457012")
                            contador_pg += 1
                            ramal = True
                case 'EI':
                    print(
                        "Iniciando processo de pagar TRE Posicão: EI")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457103)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457103")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456457106723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457106")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452122)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF 3M A 4M EX TRE COMPX C"
                                                     + " - CODIGO: 452122")

                                if contador_pg >= num_tse_linhas:
                                    break

                                if ramal is False:
                                    # Botão localizar
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457013"
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)
                                    preco.modifyCell(
                                        preco.CurrentCellRow, "QUANT", "1")
                                    preco.setCurrentCell(
                                        preco.CurrentCellRow, "QUANT")
                                    preco.pressEnter()
                                    # TRE sem fornecimento Código: 457013
                                    # TRE sem fornecimento Código: 457033
                                    print(
                                        "Pago 1 UN de TRE CER/PVC"
                                        + " 3M A 4M EI AVUL S/REP - CODIGO: 457013")
                                    contador_pg += 1
                                    ramal = True

                                # 8140 é módulo Investimento.

                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                n_etapa = preco.GetCellValue(
                                    preco.CurrentCellRow, "ETAPA")

                                if not n_etapa == operacao_rep:
                                    preco.pressToolbarButton("&FIND")
                                    session.findById(
                                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                    session.findById(
                                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(0)
                                    session.findById("wnd[1]").sendVKey(12)

                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457013"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457013
                            # TRE com fornecimento Código: 457033
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 3M A 4M EI AVUL S/REP - CODIGO: 457013")
                            contador_pg += 1
                            ramal = True
                case 'TO':
                    print(
                        "Iniciando processo de pagar TRE Posicão: TO")
                    preco = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        ramal = False
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
                                    preco_reposicao = str(457103)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457103")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456457106723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457106")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452131)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF 3M A 4M EX TRE COMPX C"
                                                     + " - CODIGO: 452131")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457014"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457014
                                # TRE com fornecimento Código: 457034
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + " 3M A 4M TO AVUL S/REP - CODIGO: 457014")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457014"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457014
                            # TRE com fornecimento Código: 457034
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 3M A 4M TO AVUL S/REP - CODIGO: 457014")
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
                                    preco_reposicao = str(457103)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP CIM  - CODIGO: 457103")
                                if pavimento[0] in dict_reposicao['especial']:
                                    preco_reposicao = str(456457106723)
                                    txt_reposicao = (
                                        "Pago 1 UN de LRP ESP  - CODIGO: 457106")
                                if pavimento[0] in dict_reposicao['asfalto_frio']:
                                    preco_reposicao = str(452131)
                                    txt_reposicao = ("Pago 1 UN de LPB ASF 3M A 4M TO TRE COMPX C"
                                                     + " - CODIGO: 452131")

                            if contador_pg >= num_tse_linhas:
                                return

                            if ramal is False:
                                # Botão localizar
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457015"
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)
                                preco.modifyCell(
                                    preco.CurrentCellRow, "QUANT", "1")
                                preco.setCurrentCell(
                                    preco.CurrentCellRow, "QUANT")
                                preco.pressEnter()
                                # TRE sem fornecimento Código: 457015
                                # TRE com fornecimento Código: 457035
                                print(
                                    "Pago 1 UN de TRE CER/PVC"
                                    + " 3M A 4M PO AVUL S/REP - CODIGO: 457015")
                                contador_pg += 1
                                ramal = True

                            # 8140 é módulo Investimento.

                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            n_etapa = preco.GetCellValue(
                                preco.CurrentCellRow, "ETAPA")

                            if not n_etapa == operacao_rep:
                                preco.pressToolbarButton("&FIND")
                                session.findById(
                                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                                session.findById(
                                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(0)
                                session.findById("wnd[1]").sendVKey(12)

                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print(txt_reposicao)
                            contador_pg += 1

                            # 5140 é módulo despesa para cimentado e especial.

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "457015"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRE sem fornecimento Código: 457015
                            # TRE com fornecimento Código: 457035
                            print(
                                "Pago 1 UN de TRE CER/PVC"
                                + " 3M A 4M PO AVUL S/REP - CODIGO: 457015")
                            contador_pg += 1
                            ramal = True
                case _:
                    return

        if posicao_rede or profundidade is None:
            return
