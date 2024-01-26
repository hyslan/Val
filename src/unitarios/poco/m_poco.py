'''Módulo Família Poço Unitário.'''
from src.lista_reposicao import dict_reposicao
from src.sap_connection import connect_to_sap
from src.unitarios.localizador import btn_localizador


class Poco:
    '''Classe Unitária de Poço.'''

    @staticmethod
    def niv_cx_parada(corte,
                      relig,
                      reposicao,
                      num_tse_linhas,
                      etapa_reposicao,
                      posicao_rede,
                      profundidade
                      ):
        '''Método Nivelamento de Caixa de Parada'''
        session = connect_to_sap()
        print(
            "Iniciando processo de pagar NIVCX REGISTRO DE PARADA - Código: 456111"
        )
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            cx_parada = False
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
                        preco_reposicao = '050401'  # em construção
                        txt_reposicao = (
                            "Pago 1 UN de LRP CIM  - CODIGO: ")
                    if pavimento[0] in dict_reposicao['especial']:
                        preco_reposicao = '050402'  # em construção
                        txt_reposicao = (
                            "Pago 1 UN de LRP ESP  - CODIGO: ")
                    if pavimento[0] in dict_reposicao['asfalto_frio']:
                        preco_reposicao = str(451123)
                        txt_reposicao = ("Pago 1 UN de LPB ASF NIV CX PARADA COMPX C"
                                         + " - CODIGO: 451123")

                    if cx_parada is False:
                        # Botão localizar
                        codigo = "456111"
                        btn_localizador(preco, session, codigo)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        print(
                            "Pago 1 UN de "
                            + "NIVCX REGISTRO DE PARADA - CODIGO: 456111")
                        contador_pg += 1
                        cx_parada = True

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

            if cx_parada is False:
                # Botão localizar
                codigo = "456111"
                btn_localizador(preco, session, codigo)
                preco.modifyCell(
                    preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(
                    preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print(
                    "Pago 1 UN de "
                    + "NIVCX REGISTRO DE PARADA - CODIGO: 456111")
                contador_pg += 1
                cx_parada = True

    @staticmethod
    def troca_de_caixa_de_parada(corte,
                                 relig,
                                 reposicao,
                                 num_tse_linhas,
                                 etapa_reposicao,
                                 posicao_rede,
                                 profundidade
                                 ):
        '''Troca de Caixa de Parada - Código 456112'''
        session = connect_to_sap()
        print(
            "Iniciando processo de pagar TROCA DE CAIXA DE PARADA - Código: 456112"
        )
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            cx_parada = False
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
                        preco_reposicao = '050401'  # em construção
                        txt_reposicao = (
                            "Pago 1 UN de LRP CIM  - CODIGO: ")
                    if pavimento[0] in dict_reposicao['especial']:
                        preco_reposicao = '050402'  # em construção
                        txt_reposicao = (
                            "Pago 1 UN de LRP ESP  - CODIGO: ")
                    if pavimento[0] in dict_reposicao['asfalto_frio']:
                        preco_reposicao = str(451123)
                        txt_reposicao = ("Pago 1 UN de LPB ASF NIV CX PARADA COMPX C"
                                         + " - CODIGO: 451123")

                    if cx_parada is False:
                        # Botão localizar
                        codigo = "456112"
                        btn_localizador(preco, session, codigo)
                        preco.modifyCell(
                            preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(
                            preco.CurrentCellRow, "QUANT")
                        preco.pressEnter()
                        print(
                            "Pago 1 UN de "
                            + "DETCX  DESC, NIV CX PARADA S/REP - CODIGO: 456112")
                        contador_pg += 1
                        cx_parada = True

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

            if cx_parada is False:
                # Botão localizar
                codigo = "456112"
                btn_localizador(preco, session, codigo)
                preco.modifyCell(
                    preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(
                    preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print(
                    "Pago 1 UN de "
                    + "DETCX  DESC, NIV CX PARADA S/REP - CODIGO: 456112")
                contador_pg += 1
                cx_parada = True

    @staticmethod
    def nivelamento(corte,
                    relig,
                    reposicao,
                    num_tse_linhas,
                    etapa_reposicao,
                    posicao_rede,
                    profundidade
                    ):
        '''Nivelamentos com e sem reposição.'''
        session = connect_to_sap()
        print(
            "Iniciando processo de pagar Nivelamentos."
        )
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            if reposicao:
                if reposicao[0] in dict_reposicao['cimentado']:
                    valor = "456207"
                    txt_valor = (
                        "Pago 1 UN de NPV C/REP PASS"
                    )

                if reposicao[0] in dict_reposicao['especial']:
                    valor = "456207"
                    txt_valor = (
                        "Pago 1 UN de NPV C/REP PASS"
                    )

                if reposicao[0] in dict_reposicao['asfalto_frio']:
                    valor = "451208"
                    txt_valor = (
                        "Pago 1 UN de NPV C/LPB LEITO  COMPX C"
                    )
                # Botão localizar
                print(f"preço: {valor}")
                preco.pressToolbarButton("&FIND")
                session.findById(
                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = valor
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]").sendVKey(12)
                preco.modifyCell(
                    preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(
                    preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print(txt_valor)
            else:
                valor = "456206"
                txt_valor = (
                    "Pago 1 UN de NPV PASS    S/REP"
                )
                # Botão localizar
                preco.pressToolbarButton("&FIND")
                session.findById(
                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = valor
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]").sendVKey(12)
                preco.modifyCell(
                    preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(
                    preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print(txt_valor)
