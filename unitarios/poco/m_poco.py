'''Módulo Família Poço Unitário.'''
from lista_reposicao import dict_reposicao
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
