# religacao.py
"""Módulo Família Religação Unitário."""
# pylint: disable=W0611
from python.src.lista_reposicao import dict_reposicao
from python.src.unitarios.localizador import btn_localizador


class Religacao:
    """Classe de Religação Unitário."""

    def __init__(self, etapa, corte, relig, reposicao, num_tse_linhas,
                 etapa_reposicao, identificador, posicao_rede,
                 profundidade, session, preco):
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
        self.preco = preco

    def restabelecida(self):
        """Método para definir de qual forma foi restabelecida e
        pagar de acordo com as informações dadas, caso contrário,
        pagar como ramal se tiver reposição ou cavalete.
        """
        try:
            if self.relig == "CAVALETE":
                print("Iniciando processo de pagar RELIG CV - Código: 456037")
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(self.preco, self.session, "456037")
                    self.preco.modifyCell(
                        self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(
                        self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    print("Pago 1 UN de RELIG CV - CODIGO: 456037")
                    return

            if self.relig in ("RAMAL PEAD", "PASSEIO") or self.reposicao:
                print(
                    "Iniciando processo de pagar RELIG RAMAL AG S/REP - Código: 456039")
                self.preco.GetCellValue(0, "NUMERO_EXT")

                ramal = False
                contador_pg = 0
                # Function lambda com list compreenhension para matriz de reposições.
                if self.reposicao:
                    rep_com_etapa = [(x, y)
                                     for x, y in zip(self.reposicao, self.etapa_reposicao, strict=False)]

                    for pavimento in rep_com_etapa:
                        operacao_rep = pavimento[1]
                        if operacao_rep == "0":
                            operacao_rep = "0010"
                        # 0 é tse da reposição;
                        # 1 é etapa da tse da reposição;
                        if pavimento[0] in dict_reposicao["cimentado"]:
                            preco_reposicao = str(456041)
                            txt_reposicao = (
                                "Pago 1 UN de LRP CIM RELIGACAO DE LIGACAO SUPR - CODIGO: 456041")
                        if pavimento[0] in dict_reposicao["especial"]:
                            preco_reposicao = str(456042)
                            txt_reposicao = (
                                "Pago 1 UN de LRP ESP RELIGACAO DE LIGACAO SUPR - CODIGO: 456042")
                        if pavimento[0] in dict_reposicao["asfalto_frio"]:
                            preco_reposicao = str(451043)
                            txt_reposicao = ("Pago 1 UN de LPB ASF SUPRE  LAG COMPX C"
                                             + " - CODIGO: 456042")

                        if contador_pg >= self.num_tse_linhas:
                            return

                        if ramal is False:
                            btn_localizador(self.preco, self.session, "456039")
                            # Marca pagar na TSE
                            self.preco.modifyCell(
                                self.preco.CurrentCellRow, "QUANT", "1")
                            self.preco.setCurrentCell(
                                self.preco.CurrentCellRow, "QUANT")
                            self.preco.pressEnter()
                            print(
                                "Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                            contador_pg += 1
                            ramal = True

                            # 660 é módulo despesa.
                            btn_localizador(
                                self.preco, self.session, preco_reposicao)
                            item_preco = self.preco.GetCellValue(
                                self.preco.CurrentCellRow, "ITEM",
                            )
                            if item_preco in ("300", "1790"):
                                self.preco.modifyCell(
                                    self.preco.CurrentCellRow, "QUANT", "1")
                                self.preco.setCurrentCell(
                                    self.preco.CurrentCellRow, "QUANT")
                                self.preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1

                        # 1820 é módulo despesa para cimentado e especial.
                        if preco_reposicao in ("456041", "456042"):
                            btn_localizador(
                                self.preco, self.session, preco_reposicao)
                            item_preco = self.preco.GetCellValue(
                                self.preco.CurrentCellRow, "ITEM",
                            )
                            if item_preco in ("1820", "1830", "670",
                                              "5310"):
                                self.preco.modifyCell(
                                    self.preco.CurrentCellRow, "QUANT", "1")
                                self.preco.setCurrentCell(
                                    self.preco.CurrentCellRow, "QUANT")
                                self.preco.pressEnter()
                                print(txt_reposicao)
                                contador_pg += 1
                    return

                if ramal is False:
                    contador_pg = 0
                    btn_localizador(
                        self.preco, self.session, "456039")
                    self.preco.modifyCell(
                        self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(
                        self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    print(
                        "Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                    contador_pg += 1
                    ramal = True
                    return

            if self.relig in ("FERRULE", "TOMADA/FERRULE"):
                print(
                    "Iniciando processo de pagar RELIG  TMD AG  S/REP - Código: 456040")
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(
                        self.preco, self.session, "456040")
                    self.preco.modifyCell(
                        self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(
                        self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    print(
                        "Pago 1 UN de RELIG  TMD AG  S/REP - CODIGO: 456040")
                    return

            if self.relig is None:
                print("Religação não informada. \n Pagando como RELIG CV.")
                self.preco.GetCellValue(0, "NUMERO_EXT")
                if self.preco is not None:
                    btn_localizador(
                        self.preco, self.session, "456037")
                    self.preco.modifyCell(
                        self.preco.CurrentCellRow, "QUANT", "1")
                    self.preco.setCurrentCell(
                        self.preco.CurrentCellRow, "QUANT")
                    self.preco.pressEnter()
                    print("Pago 1 UN de RELIG CV - CODIGO: 456037")
                    return

        except Exception as erro:
            print(f"Erro ao pagar religação: {erro}")

        # Confirmação da precificação.
        self.preco.pressEnter()
