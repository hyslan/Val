# hidrometro_material.py
'''Módulo dos materiais de família Rede de Esgoto.'''
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from wms import testa_material_sap
from wms import materiais_contratada


session = connect_to_sap()
(
    lista,
    _,
    _,
    _,
    planilha,
    _,
    _,
    _,
    _,
    _,
    tb_contratada,
    _,
    *_,
) = load_worksheets()


class RedeEsgotoMaterial:
    '''Classe de materiais de CRE.'''
    tamanhos_junta_luva = [
        "30002958",
        "30005617",
        "30000357"
    ]
    tamanho_junta_adap = [
        "30005615",
        "30001528",
        "30005619",
        "30001529"
    ]

    def __init__(self, int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais) -> None:
        self.int_num_lordem = int_num_lordem
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais

    def receita_reparo_de_rede_de_esgoto(self):
        '''Padrão de materiais na classe CRE.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        match self.diametro_rede:
            case '100':
                junta_esgoto = "30002958"
                junta_esgoto_adap = "30005615"
            case '150':
                junta_esgoto = "30005617"
                junta_esgoto_adap = "30001528"
            case '200':
                junta_esgoto = "30000357"
                junta_esgoto_adap = "30005619"
            case '300':
                junta_esgoto_adap = "30001529"
            case _:
                print("Diâmetro da rede não informado, aplicando Junta DN 150.")
                junta_esgoto = "30005617"
                junta_esgoto_adap = "30001528"
        if sap_material is None:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona JUNTA FLEX ESGOTO LUVA
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", junta_esgoto
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1

        else:
            material_lista = []
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas

            # Loop de Lista Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(
                    n_material, "ETAPA")
                material_lista.append(sap_material)

            # Loop busca junta flex adap correta.
            for junta_ad in RedeEsgotoMaterial.tamanho_junta_adap:
                if junta_ad in material_lista:
                    print(junta_ad)
                    if junta_ad == junta_esgoto_adap:
                        print("Junta correta.")
                    else:
                        print("Colocando a junta correta.")
                        for n_material in range(num_material_linhas):
                            # Pega valor da célula 0
                            sap_material = self.tb_materiais.GetCellValue(
                                n_material, "MATERIAL")
                            sap_etapa_material = self.tb_materiais.GetCellValue(
                                n_material, "ETAPA")

                            if sap_material == junta_ad:
                                self.tb_materiais.modifyCheckbox(
                                    n_material, "ELIMINADO", True
                                )
                                self.tb_materiais.InsertRows(
                                    str(ultima_linha_material))
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "ETAPA", sap_etapa_material
                                )
                                # Adiciona a Junta Flex Adap correta..
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "MATERIAL", junta_esgoto_adap
                                )
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "QUANT", "1"
                                )
                                self.tb_materiais.setCurrentCell(
                                    ultima_linha_material, "QUANT"
                                )
                                ultima_linha_material = ultima_linha_material + 1

            # Loop busca junta flex luva correta.
            for junta_luva in RedeEsgotoMaterial.tamanhos_junta_luva:
                if junta_luva in material_lista:
                    if junta_luva == junta_esgoto:
                        print("Junta correta.")
                    else:
                        print("Colocando a junta correta.")
                        for n_material in range(num_material_linhas):
                            # Pega valor da célula 0
                            sap_material = self.tb_materiais.GetCellValue(
                                n_material, "MATERIAL")
                            sap_etapa_material = self.tb_materiais.GetCellValue(
                                n_material, "ETAPA")

                            if sap_material == junta_luva:
                                self.tb_materiais.modifyCheckbox(
                                    n_material, "ELIMINADO", True
                                )
                                self.tb_materiais.InsertRows(
                                    str(ultima_linha_material))
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "ETAPA", sap_etapa_material
                                )
                                # Adiciona a Junta Flex correta..
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "MATERIAL", junta_esgoto
                                )
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "QUANT", "1"
                                )
                                self.tb_materiais.setCurrentCell(
                                    ultima_linha_material, "QUANT"
                                )
                                ultima_linha_material = ultima_linha_material + 1

            # Área geral de materiais
            n_material = 0
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")

                if sap_material == '30022469':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona Curva 45G ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30005282"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                if sap_material == '30005894':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona LUVA CORRER BB ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30002797"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                if sap_material == '30005283':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona Curva 45G ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30005282"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais)

    def receita_reparo_de_ramal_de_esgoto(self):
        '''Padrão de materiais na classe Ramal de Esgoto.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        match self.diametro_ramal:
            case 'DN_100':
                junta_esgoto = "30002958"
                junta_esgoto_adap = "30005615"
            case 'DN_150':
                junta_esgoto = "30005617"
                junta_esgoto_adap = "30001528"
            case 'DN_200':
                junta_esgoto = "30000357"
                junta_esgoto_adap = "30005619"
            case 'DN_300':
                junta_esgoto_adap = "30001529"
            case _:
                print("Diâmetro do ramal não informado, aplicando Junta DN 100.")
                junta_esgoto = "30002958"
                junta_esgoto_adap = "30005615"
        if sap_material is None:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona JUNTA FLEX ESGOTO LUVA
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", junta_esgoto
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1

        else:
            material_lista = []
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas
            # Loop de Lista Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                material_lista.append(sap_material)

            # Loop busca junta flex adap correta.
            for junta_ad in RedeEsgotoMaterial.tamanho_junta_adap:
                if junta_ad in material_lista:
                    if junta_ad == junta_esgoto_adap:
                        print("Junta correta.")
                    else:
                        print("Colocando a junta correta.")
                        for n_material in range(num_material_linhas):
                            # Pega valor da célula 0
                            sap_material = self.tb_materiais.GetCellValue(
                                n_material, "MATERIAL")
                            sap_etapa_material = self.tb_materiais.GetCellValue(
                                n_material, "ETAPA")

                            if sap_material == junta_ad:
                                self.tb_materiais.modifyCheckbox(
                                    n_material, "ELIMINADO", True
                                )
                                self.tb_materiais.InsertRows(
                                    str(ultima_linha_material))
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "ETAPA", sap_etapa_material
                                )
                                # Adiciona a Junta Flex correta..
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "MATERIAL", junta_esgoto_adap
                                )
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "QUANT", "1"
                                )
                                self.tb_materiais.setCurrentCell(
                                    ultima_linha_material, "QUANT"
                                )
                                ultima_linha_material = ultima_linha_material + 1

            # Loop busca junta flex luva correta.
            for junta_luva in RedeEsgotoMaterial.tamanhos_junta_luva:
                if junta_luva in material_lista:
                    if junta_luva == junta_esgoto:
                        print("Junta correta.")
                    else:
                        print("Colocando a junta correta.")
                        for n_material in range(num_material_linhas):
                            # Pega valor da célula 0
                            sap_material = self.tb_materiais.GetCellValue(
                                n_material, "MATERIAL")
                            sap_etapa_material = self.tb_materiais.GetCellValue(
                                n_material, "ETAPA")
                            if sap_material == junta_luva:
                                self.tb_materiais.modifyCheckbox(
                                    n_material, "ELIMINADO", True
                                )
                                self.tb_materiais.InsertRows(
                                    str(ultima_linha_material))
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "ETAPA", sap_etapa_material
                                )
                                # Adiciona a Junta Flex correta..
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "MATERIAL", junta_esgoto
                                )
                                self.tb_materiais.modifyCell(
                                    ultima_linha_material, "QUANT", "1"
                                )
                                self.tb_materiais.setCurrentCell(
                                    ultima_linha_material, "QUANT"
                                )
                                ultima_linha_material = ultima_linha_material + 1

            # Área Geral dos Materiais
            n_material = 0
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(
                    n_material, "ETAPA")

                if sap_material == '30022469':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", sap_etapa_material
                    )
                    # Adiciona Curva 45G ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30005282"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                if sap_material == '30000139':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona Curva 90G ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30022735"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                # O ventilador bugado do SAP.
                if sap_material == '30002858':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

                if sap_material == '30005616':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", sap_etapa_material
                    )
                    # Adiciona Curva 90G ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30005617"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                if sap_material == '30005894':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona LUVA CORRER BB ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30002797"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                if sap_material == '30005283':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona Curva 45G ESG DN 100 vigente.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30005282"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1

                # Removendo JUNTA FLEX ESGOTO ADAP DN 100MMX100MM porque acabou.
                if sap_material == '30005615':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais)
