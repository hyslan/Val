'''Módulo Família Ligação Água Unitário.'''
# pylint: disable=W0611
import sys
from lista_reposicao import dict_reposicao
from sap_connection import connect_to_sap
session = connect_to_sap()


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
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        n_etapa = preco.GetCellValue(
                            preco.CurrentCellRow, "ETAPA")

                        if not n_etapa == operacao_rep:
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
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
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        n_etapa = preco.GetCellValue(
                            preco.CurrentCellRow, "ETAPA")

                        if not n_etapa == operacao_rep:
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
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

        if posicao_rede is None:
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
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        n_etapa = preco.GetCellValue(
                            preco.CurrentCellRow, "ETAPA")

                        if not n_etapa == operacao_rep:
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
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
                        session.findById("wnd[1]").sendVKey(0)
                        session.findById("wnd[1]").sendVKey(12)
                        n_etapa = preco.GetCellValue(
                            preco.CurrentCellRow, "ETAPA")

                        if not n_etapa == operacao_rep:
                            preco.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
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

        if posicao_rede is None:
            print("Sem posição de rede informada")
            return
