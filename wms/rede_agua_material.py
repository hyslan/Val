# hidrometro_material.py
'''Módulo dos materiais de família Rede de Água.'''
from excel_tbs import load_worksheets
from wms import testa_material_sap
from wms import materiais_contratada


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
    tb_contratada_gb,
    _,
    *_,
) = load_worksheets()


class RedeAguaMaterial:
    '''Classe de materiais de CRA.'''

    def __init__(self, int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais,
                 contrato) -> None:
        self.int_num_lordem = int_num_lordem
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato

    def materiais_vigentes(self):
        '''Materiais com estoque.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        if sap_material is not None:
            num_material_linhas = self.tb_materiais.RowCount
            n_material = 0
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(
                    n_material, "ETAPA")

                match sap_material:
                    case '30029526':
                        if self.contrato == '4600041302':
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True
                            )

                    # case '30001122':
                    #     self.tb_materiais.modifyCheckbox(
                    #         n_material, "ELIMINADO", True
                    #     )

                    case '30002735':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30011136':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30005088':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", sap_etapa_material
                        )
                        # Adiciona CONEXÕES METALICAS COTOVELO FEMEA DN 3/4.
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "MATERIAL", "30002394"
                        )
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "QUANT", "1"
                        )
                        self.tb_materiais.setCurrentCell(
                            ultima_linha_material, "QUANT"
                        )
                        ultima_linha_material = ultima_linha_material + 1

                    case '30004097':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30002152':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30002151':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30002802':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30004104':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30007931':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

    def receita_reparo_de_rede_de_agua(self):
        '''Padrão de materiais na classe CRA.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        abracadeira_dn75 = False
        tubo_pba_dn50 = False
        if sap_material is None:
            print("sem material.")
        else:
            material_lista = []
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                material_lista.append(sap_material)
            n_material = 0
            ultima_linha_material = num_material_linhas
            if "30008103" in material_lista:
                abracadeira_dn75 = True
            if "30028862" in material_lista:
                tubo_pba_dn50 = True
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")

                if sap_material in ('30004097', '30002152', '30002151') \
                        and abracadeira_dn75 is False and self.diametro_rede == '100':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona ABRACADEIRA FF REPARO TUBO DN100 LMIN=150.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30008103"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1
                    abracadeira_dn75 = True

            if tubo_pba_dn50 is False and self.diametro_rede == '50':
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", self.identificador[1]
                )
                # Adiciona TUBO PBA DN 50 1,00 MPA JEI/JERI CM 6M.
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "30028862"
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1"
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT"
                )
                ultima_linha_material = ultima_linha_material + 1

            self.materiais_vigentes()
            # Materiais do Global.
            materiais_contratada.materiais_contratada(
                self.tb_materiais, self.contrato)

    def receita_troca_de_conexao_de_ligacao_de_agua(self):
        '''Padrão de materiais na classe Troca de Conexão de Ligação de Água.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        if sap_material is None:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona CONEXOES MET LIGACOES FEMEA DN 20.
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "30002394"
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1

            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona REGISTRO METALICO RAMAL PREDIAL DN 20.
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "30006747"
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

            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                material_lista.append(sap_material)

            if '30002394' not in material_lista:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", self.identificador[1]
                )
                # Adiciona CONEXOES MET LIGACOES FEMEA DN 20.
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "30002394"
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1"
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT"
                )
                ultima_linha_material = ultima_linha_material + 1

            if '30006747' not in material_lista:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", self.identificador[1]
                )
                # Adiciona REGISTRO METALICO RAMAL PREDIAL DN 20.
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "30006747"
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1"
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT"
                )
                ultima_linha_material = ultima_linha_material + 1

            self.materiais_vigentes()
            # Materiais do Global.
            materiais_contratada.materiais_contratada(
                self.tb_materiais, self.contrato)

    def receita_reparo_de_ramal_de_agua(self):
        '''Padrão de materiais no reparo de Ligação de Água.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        if sap_material is None:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona TUDO PEAD DN 20.
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "30001848"
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

            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                material_lista.append(sap_material)

            if '30001848' not in material_lista:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", self.identificador[1]
                )
                # Adiciona TUBO PEAD DN 20 - PE 80 - 1.0 MPA.
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "30001848"
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1"
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT"
                )
                ultima_linha_material = ultima_linha_material + 1

            self.materiais_vigentes()
            # Materiais do Global.
            materiais_contratada.materiais_contratada(
                self.tb_materiais, self.contrato)
