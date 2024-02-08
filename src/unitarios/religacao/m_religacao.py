# religacao.py
'''Módulo Família Religação Unitário.'''
# pylint: disable=W0611
from src.lista_reposicao import dict_reposicao
from src.unitarios.localizador import btn_localizador


class Religacao:
    '''Classe de Religação Unitário.'''

    def __init__(self, etapa, corte, relig, reposicao, num_tse_linhas,
                 etapa_reposicao, identificador, posicao_rede, profundidade, session):
        self.etapa = etapa
        self.corte = corte
        self.relig = relig
        self.reposicao = reposicao
        self.num_tse_linhas = num_tse_linhas
        self.etapa_reposicao = etapa_reposicao
        self.posicao_rede = posicao_rede
        self.profundidade = profundidade
        self.session = session
        self.identificador = identificador

    def restabelecida(self):
        '''Método para definir de qual forma foi restabelecida e 
        pagar de acordo com as informações dadas, caso contrário,
        pagar como ramal se tiver reposição ou cavalete.'''
        if self.relig == 'CAVALETE':
            print("Iniciando processo de pagar RELIG CV - Código: 456037")
            preco = self.session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
            preco.GetCellValue(0, "NUMERO_EXT")
            if preco is not None:
                btn_localizador(preco, self.session, "456037")
                preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print("Pago 1 UN de RELIG CV - CODIGO: 456037")

        if self.relig == 'RAMAL PEAD' or self.reposicao:
            print(
                "Iniciando processo de pagar RELIG RAMAL AG S/REP - Código: 456039")
            preco = self.session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
            preco.GetCellValue(0, "NUMERO_EXT")

            ramal = False
            contador_pg = 0
            # Function lambda com list compreenhension para matriz de reposições.
            if self.reposicao:
                rep_com_etapa = [(x, y)
                                 for x, y in zip(self.reposicao, self.etapa_reposicao)]

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

                    if contador_pg >= self.num_tse_linhas:
                        return

                    if ramal is False:
                        btn_localizador(preco, self.session, "456039")
                        # Marca pagar na TSE
                        preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                        preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                        print(
                            "Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                        contador_pg += 1
                        ramal = True

                        # 660 é módulo despesa.
                        btn_localizador(preco, self.session, preco_reposicao)
                        item_preco = preco.GetCellValue(
                            preco.CurrentCellRow, "ITEM"
                        )
                        if item_preco == '300':
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                            print(txt_reposicao)
                            contador_pg += 1

                    # 1820 é módulo despesa para cimentado e especial.
                    if preco_reposicao in ('456041', '456042'):
                        btn_localizador(
                            preco, self.session, preco_reposicao)
                        item_preco = preco.GetCellValue(
                            preco.CurrentCellRow, "ITEM"
                        )
                        if item_preco == '1820':
                            preco.modifyCell(
                                preco.CurrentCellRow, "QUANT", "1")
                            preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                            print(txt_reposicao)
                            contador_pg += 1

            if ramal is False:
                contador_pg = 0
                btn_localizador(
                    preco, self.session, "456039")
                preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                print(
                    "Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                contador_pg += 1
                ramal = True

        if self.relig == 'FERRULE':
            print(
                "Iniciando processo de pagar RELIG  TMD AG  S/REP - Código: 456040")
            preco = self.session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
            preco.GetCellValue(0, "NUMERO_EXT")
            if preco is not None:
                btn_localizador(
                    preco, self.session, "456040")
                preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print(
                    "Pago 1 UN de RELIG  TMD AG  S/REP - CODIGO: 456040")

        else:
            print("Religação não informada. \n Pagando como RELIG CV.")
            preco = self.session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
            preco.GetCellValue(0, "NUMERO_EXT")
            if preco is not None:
                btn_localizador(
                    preco, self.session, "456037")
                preco.modifyCell(preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print("Pago 1 UN de RELIG CV - CODIGO: 456037")

        # Confirmação da precificação.
        preco.pressEnter()
