'''Módulo Família Ligação Água Unitário.'''
# pylint: disable=W0611
import sys
from src.lista_reposicao import dict_reposicao
from src.sap_connection import connect_to_sap


class LigacaoAgua:
    '''Classe de Ligação Esgoto Unitário.'''
    @staticmethod
    def ligacao_agua(corte,
                     relig,
                     reposicao,
                     num_tse_linhas,
                     etapa_reposicao,
                     posicao_rede,
                     profundidade
                     ):
        '''Método para definir de qual forma foi a Ligação de água e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a L.A .'''
        session = connect_to_sap()

        if posicao_rede == 'PA':
            print(
                "Iniciando processo de pagar ligação de água s/v Posição: PA"
            )
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
                            preco_reposicao = str(456471)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM  - CODIGO: 456471")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456472)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP  - CODIGO: 456472")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451495)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451495")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456451"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456451
                            # LESG com fornecimento Código: 456461
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "PA AVUL S/REP - CODIGO: 456451")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456451"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # LESG sem fornecimento Código: 456451
                    # LESG com fornecimento Código: 456461
                    print(
                        "Pago 1 UN de LAG ATE 32MM"
                        + "PA AVUL S/REP - CODIGO: 456451")
                    contador_pg += 1
                    ramal = True

        # MND
        if posicao_rede in ('TA', 'EI', 'TO', 'PO'):
            print(
                "Iniciando processo de pagar ligação de água s/v Posição: PA"
            )
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
                            preco_reposicao = str(456493)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM MND  - CODIGO: 456493")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456493)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP MND  - CODIGO: 456493")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451495)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451495")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456491"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456491
                            # LESG com fornecimento Código: 456492
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "MND AVUL S/REP - CODIGO: 456491")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456491"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # LESG sem fornecimento Código: 456491
                    # LESG com fornecimento Código: 456492
                    print(
                        "Pago 1 UN de LAG ATE 32MM"
                        + "MND AVUL S/REP - CODIGO: 456491")
                    contador_pg += 1
                    ramal = True

        if not posicao_rede:
            print("Sem posição de rede informada")
            return

    @staticmethod
    def tra_nv(corte,
               relig,
               reposicao,
               num_tse_linhas,
               etapa_reposicao,
               posicao_rede,
               profundidade
               ):
        '''Método para definir de qual forma foi a TRA não visível e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a TRA NV .'''
        session = connect_to_sap()

        if posicao_rede == 'PA':
            print(
                "Iniciando processo de pagar TRA NV Posição: PA"
            )
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
                            preco_reposicao = str(456471)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM  - CODIGO: 456471")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456822)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP  - CODIGO: 456822")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451885)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451885")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456801"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456801
                            # LESG com fornecimento Código: 456811
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "PA AVUL S/REP - CODIGO: 456801")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456801"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # LESG sem fornecimento Código: 456801
                    # LESG com fornecimento Código: 456811
                    print(
                        "Pago 1 UN de TRA ATE 32MM"
                        + "CORR PA AVUL S/REP - CODIGO: 456801")
                    contador_pg += 1
                    ramal = True

        # MND
        if posicao_rede in ('TA', 'EI', 'TO', 'PO'):
            print(
                "Iniciando processo de pagar ligação de água s/v Posição: PA"
            )
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
                            preco_reposicao = str(456883)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM MND  - CODIGO: 456883")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456884)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP MND  - CODIGO: 456884")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451885)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451885")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456881"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456881
                            # LESG com fornecimento Código: 456882
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + " MND AVUL S/REP - CODIGO: 456491")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456881"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # LESG sem fornecimento Código: 456881
                    # LESG com fornecimento Código: 456882
                    print(
                        "Pago 1 UN de TRA ATE 32MM"
                        + " CORR MND AVUL S/REP - CODIGO: 456881")
                    contador_pg += 1
                    ramal = True

        if not posicao_rede:
            print("Sem posição de rede informada")
            return

    @staticmethod
    def png(corte,
            relig,
            reposicao,
            num_tse_linhas,
            etapa_reposicao,
            posicao_rede,
            profundidade
            ):
        '''Método para definir de qual forma foi a PNG e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a PNG.'''
        session = connect_to_sap()

        if posicao_rede == 'PA':
            print(
                "Iniciando processo de pagar PNG Água Posição: PA"
            )
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
                            preco_reposicao = str(456391)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM  - CODIGO: 456391")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456392)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP  - CODIGO: 456392")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451415)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451415")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456371"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # PNG sem fornecimento Código: 456371
                            # PNG com fornecimento Código: 456381
                            print(
                                "Pago 1 UN de LAG"
                                + "PA SUCES S/REP - CODIGO: 456801")
                            contador_pg += 1
                            ramal = True

                        # 3980 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456371"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # PNG sem fornecimento Código: 456371
                    # PNG com fornecimento Código: 456381
                    print(
                        "Pago 1 UN de LAG"
                        + "CORR PA SUCES S/REP - CODIGO: 456371")
                    contador_pg += 1
                    ramal = True

        # MND
        if posicao_rede in ('TA', 'EI', 'TO', 'PO'):
            print(
                "Iniciando processo de pagar PNG Água Posição: PA"
            )
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
                            preco_reposicao = str(456413)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM MND  - CODIGO: 456413")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456414)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP MND  - CODIGO: 456414")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451415)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451415")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456411"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # PNG sem fornecimento Código: 456411
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + " MND AVUL S/REP - CODIGO: 456411")
                            contador_pg += 1
                            ramal = True

                        # 4100 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456411"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # PNG sem fornecimento Código: 456411
                    print(
                        "Pago 1 UN de LAG"
                        + " MND SUCES S/REP - CODIGO: 456411")
                    contador_pg += 1
                    ramal = True

        if not posicao_rede:
            print("Sem posição de rede informada")
            return

    @staticmethod
    def subst_agua(corte,
                   relig,
                   reposicao,
                   num_tse_linhas,
                   etapa_reposicao,
                   posicao_rede,
                   profundidade
                   ):
        '''Método para definir de qual forma foi a Substituição de 
        Ligação de água e pagar de acordo com as
        informações dadas, caso contrário,
        não pagar a SUBST L.A .'''
        session = connect_to_sap()

        if posicao_rede == 'PA':
            print(
                "Iniciando processo de pagar Subst. de ligação de água Posição: PA"
            )
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
                            preco_reposicao = str(456471)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM  - CODIGO: 456471")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456472)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP  - CODIGO: 456472")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451495)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451495")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456451"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456451
                            # LESG com fornecimento Código: 456461
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "PA AVUL S/REP - CODIGO: 456451")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

                            # Adicional de Supressão do FERRULE
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456506"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de ADC  SUPR TMD AG  S/REP")

                        # Etapa Pavimentos
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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456451"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # LESG sem fornecimento Código: 456451
                    # LESG com fornecimento Código: 456461
                    print(
                        "Pago 1 UN de LAG ATE 32MM"
                        + "PA AVUL S/REP - CODIGO: 456451")
                    contador_pg += 1
                    ramal = True
                    # Adicional de Supressão do FERRULE
                    preco.pressToolbarButton("&FIND")
                    session.findById(
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456506"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de ADC  SUPR TMD AG  S/REP")

        # MND
        if posicao_rede in ('TA', 'EI', 'TO', 'PO'):
            print(
                "Iniciando processo de pagar Subst. de ligação de água Posição: MND"
            )
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
                            preco_reposicao = str(456493)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM MND  - CODIGO: 456493")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456493)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP MND  - CODIGO: 456493")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451495)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                             + " - CODIGO: 451495")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456491"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # LESG sem fornecimento Código: 456491
                            # LESG com fornecimento Código: 456492
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "MND AVUL S/REP - CODIGO: 456491")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

                            # Adicional de Supressão do FERRULE
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456506"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de ADC  SUPR TMD AG  S/REP")

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456491"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # LESG sem fornecimento Código: 456491
                    # LESG com fornecimento Código: 456492
                    print(
                        "Pago 1 UN de LAG ATE 32MM"
                        + "MND AVUL S/REP - CODIGO: 456491")
                    contador_pg += 1
                    ramal = True
                    # Adicional de Supressão do FERRULE
                    preco.pressToolbarButton("&FIND")
                    session.findById(
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456506"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de ADC  SUPR TMD AG  S/REP")

        if not posicao_rede:
            print("Sem posição de rede informada")
            return

    @staticmethod
    def tra_prev(corte,
                 relig,
                 reposicao,
                 num_tse_linhas,
                 etapa_reposicao,
                 posicao_rede,
                 profundidade
                 ):
        '''Método para definir de qual forma foi a TRA Preventiva e 
        pagar de acordo com as informações dadas, caso contrário,
        não pagar a TRA PREV .'''
        session = connect_to_sap()

        if posicao_rede == 'PA':
            print(
                "Iniciando processo de pagar TRA PREV Posição: PA"
            )
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
                            preco_reposicao = str(456861)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM  - CODIGO: 456861")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456862)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP  - CODIGO: 456862")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451895)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG TRA PREV COMPX C"
                                             + " - CODIGO: 451895")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456841"
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRAPREV sem fornecimento Código: 456841
                            # TRAPREV com fornecimento Código: 456851
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "PREV PA AVUL S/REP - CODIGO: 456841")
                            contador_pg += 1
                            ramal = True

                        # 2780 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456841"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # TRAPREV sem fornecimento Código: 456841
                    # TRAPREV com fornecimento Código: 456851
                    print(
                        "Pago 1 UN de TRA ATE 32MM"
                        + "PREV PA S/REP - CODIGO: 456841")
                    contador_pg += 1
                    ramal = True

        # MND
        if posicao_rede in ('TA', 'EI', 'TO', 'PO'):
            print(
                "Iniciando processo de pagar ligação de água s/v Posição: MND"
            )
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
                            preco_reposicao = str(456893)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM MND  - CODIGO: 456893")
                        if pavimento[0] in dict_reposicao['especial']:
                            preco_reposicao = str(456894)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP MND  - CODIGO: 456894")
                        if pavimento[0] in dict_reposicao['asfalto_frio']:
                            preco_reposicao = str(451895)
                            txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG TRA PREV COMPX C"
                                             + " - CODIGO: 451895")

                        if ramal is False:
                            # Botão localizar
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456891"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(
                                preco.CurrentCellRow, "QUANT")
                            preco.pressEnter()
                            # TRA PREV sem fornecimento Código: 456891
                            # TRA PREV com fornecimento Código: 456892
                            print(
                                "Pago 1 UN de LAG ATE 32MM"
                                + "PREV MND S/REP - CODIGO: 456891")
                            contador_pg += 1
                            ramal = True

                        # 4220 é módulo Investimento.

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
                        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = "456891"
                    session.findById(
                        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                    session.findById("wnd[1]").sendVKey(0)
                    session.findById("wnd[1]").sendVKey(12)
                    preco.modifyCell(
                        preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(
                        preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    # TRA PREV sem fornecimento Código: 456891
                    # TRA PREV com fornecimento Código: 456892
                    print(
                        "Pago 1 UN de TRA ATE 32MM"
                        + " TREV MND S/REP - CODIGO: 456891")
                    contador_pg += 1
                    ramal = True
