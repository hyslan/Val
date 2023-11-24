'''Módulo Unitário de cavalete'''
# cavalete.py
from sap_connection import connect_to_sap
from unitarios.localizador import btn_localizador


class Cavalete:
    '''Classe Cavalete unitário'''

    @staticmethod
    def instalado_lacre(*_):
        '''Método Instalado Lacre Diversos'''
        session = connect_to_sap()
        print("Iniciando processo de pagar LACRHD - Código: 456021")
        codigo = "456021"
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, session, codigo)
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de LACRHD - CODIGO: 456021")

    @staticmethod
    def troca_pe_cv_prev(*_):
        '''Método Troca de Pé de Cavalete Preventivo'''
        session = connect_to_sap()
        print("Iniciando processo de pagar ADC  TRC PREV PE CV - Código: 456856")
        codigo = "456856"
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            btn_localizador(preco, session, codigo)
            preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print("Pago 1 UN de ADC  TRC PREV PE CV - CODIGO: 456856")

    @staticmethod
    def troca_cv_kit(*_):
        '''Método Troca de cavalete KIT'''
        session = connect_to_sap()
        print("Iniciando processo de pagar TROCA DE CAVALETE (KIT)  - Código: 456011")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
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

    @staticmethod
    def troca_cv_por_uma(corte,
                         relig,
                         reposicao,
                         num_tse_linhas,
                         etapa_reposicao,
                         posicao_rede,
                         profundidade
                         ):
        '''Método Troca de Cavalete Preventivo'''
        session = connect_to_sap()
        print("Iniciando processo de pagar SUBST CV POR UMA")
        preco = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            if reposicao:
                codigo = "457229"
                btn_localizador(preco, session, codigo)
                item_preco = preco.GetCellValue(
                    preco.CurrentCellRow, "ITEM"
                )
                if item_preco == '300':
                    preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    print(
                        "Pago 1 UN de SUBST CV POR UMA C/REP PASS - CODIGO: 457229")

            else:
                codigo = "457230"
                btn_localizador(preco, session, codigo)
                item_preco = preco.GetCellValue(
                    preco.CurrentCellRow, "ITEM"
                )
                if item_preco == '300':
                    preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                    preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                    preco.pressEnter()
                    print(
                        "Pago 1 UN de SUBST CV POR UMA S/REP - CODIGO: 457230")
