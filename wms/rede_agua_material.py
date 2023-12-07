# hidrometro_material.py
'''Módulo dos materiais de família Rede de Água.'''
from wms import testa_material_sap
from wms import materiais_contratada
from wms import localiza_material
from sap_connection import connect_to_sap


class RedeAguaMaterial:
    '''Classe de materiais de CRA.'''

    def __init__(self, int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais,
                 contrato,
                 estoque,
                 df_materiais,
                 posicao_rede) -> None:
        self.int_num_lordem = int_num_lordem
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato
        self.estoque = estoque
        self.df_materiais = df_materiais
        self.posicao_rede = posicao_rede

    def materiais_vigentes(self):
        '''Materiais com estoque.'''
        session = connect_to_sap()
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        # CONEXOES MET LIGACOES  FÊMEA DN 20
        con_met_femea_estoque = self.estoque[self.estoque['Material'] == '30002394']
        # CONEXOES MET ADAP MACHO DN 20
        con_met_macho_estoque = self.estoque[self.estoque['Material'] == '30001346']
        # DISPOSITIVO MED PLASTICO DN 20
        disp_med_plastico_estoque = self.estoque[self.estoque['Material'] == '50000178']
        # REGISTRO METALICO RAMAL PREDIAL DN 20
        reg_met_predial_dn20_estoque = self.estoque[self.estoque['Material'] == '30006747']
        # COLAR TOM P/TUBO PE DE 32XDN 20 TE INTEG
        colar_tom_tubo_pe_dn32xdn20_estoque = self.estoque[
            self.estoque['Material'] == '30000287']
        # COLAR TOMADA ACO INOX DN50A150 X DNR20
        colar_tom_aco_inox__dn50a150xdnr20_estoque = self.estoque[
            self.estoque['Material'] == '30004702']
        # COLAR TOMADA ACO INOX DN200A300 X DNR20
        colar_tom_aco_inox__dn200a300xdnr20_estoque = self.estoque[
            self.estoque['Material'] == '30004701']
        # COLAR TOMADA FF C.INOX DN200A300 X DNR20
        colar_tom_ff_cinox__dn200a300xdnr20_estoque = self.estoque[
            self.estoque['Material'] == '30004703']
        # ABRACADEIRA FF REPARO TUBO DN75 LMIN=150
        abrac_ff_reparo_dn75_lmin150_estoque = self.estoque[self.estoque['Material'] == '30001122']
        # ABRACADEIRA INOX REPARO TUBO DN50 L=300
        abrac_inox_reparo_dn50_l300_estoque = self.estoque[self.estoque['Material'] == '30002151']
        # TUBO PEAD DN 20
        tubo_pead_dn20_estoque = self.estoque[self.estoque['Material'] == '30001848']

        if sap_material is not None:
            num_material_linhas = self.tb_materiais.RowCount
            ultima_linha_material = num_material_linhas
            tubo_pead_dn20 = self.df_materiais[self.df_materiais['Material'] == '30001848']
            con_met_femea = self.df_materiais[self.df_materiais['Material'] == '30002394']
            con_met_macho = self.df_materiais[self.df_materiais['Material'] == '30001346']
            disp_med_plastico = self.df_materiais[self.df_materiais['Material'] == '50000178']
            reg_met_predial_dn20 = self.df_materiais[self.df_materiais['Material'] == '30006747']
            colar_tom_tubo_pe_dn32xdn20 = self.df_materiais[
                self.df_materiais['Material'] == '30000287']
            colar_tom_aco_inox__dn50a150xdnr20 = self.df_materiais[
                self.df_materiais['Material'] == '30004702']
            colar_tom_aco_inox__dn200a300xdnr20 = self.df_materiais[
                self.df_materiais['Material'] == '30004701']
            colar_tom_ff_cinox__dn200a300xdnr20 = self.df_materiais[
                self.df_materiais['Material'] == '30004703']

            for material in self.df_materiais['Material']:
                match material:
                    case '30002394':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30002394')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '30001346':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30001346')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '50000178':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '50000178')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '30006747':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30006747')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '30000287':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30000287')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '30004702':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30004702')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '30004701':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30004701')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case '30004703':
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, '30004703')
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    # TUBO PEAD DN20
                    case '30001848':
                        codigo = '30001848'
                        match self.posicao_rede:
                            case 'PA':
                                resultado = localiza_material.qtd_max(
                                    material, self.estoque, 2.5, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(
                                        self.tb_materiais, session, codigo)
                                    localiza_material.qtd_correta(
                                        self.tb_materiais, "2")
                            case 'TA':
                                resultado = localiza_material.qtd_max(
                                    material, self.estoque, 4.5, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(
                                        self.tb_materiais, session, codigo)
                                    localiza_material.qtd_correta(
                                        self.tb_materiais, "4")
                            case 'EX':
                                resultado = localiza_material.qtd_max(
                                    material, self.estoque, 10, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(
                                        self.tb_materiais, session, codigo)
                                    localiza_material.qtd_correta(
                                        self.tb_materiais, "5")
                            case 'TO':
                                resultado = localiza_material.qtd_max(
                                    material, self.estoque, 15, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(
                                        self.tb_materiais, session, codigo)
                                    localiza_material.qtd_correta(
                                        self.tb_materiais, "7")
                            case 'PO':
                                resultado = localiza_material.qtd_max(
                                    material, self.estoque, 15, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(
                                        self.tb_materiais, session, codigo)
                                    localiza_material.qtd_correta(
                                        self.tb_materiais, "10")

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

                    case '30001122':
                        if abrac_ff_reparo_dn75_lmin150_estoque.empty:
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True
                            )

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
                        if not con_met_femea_estoque.empty:
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
                        if abrac_inox_reparo_dn50_l300_estoque.empty:
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

                    case '30002202':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        if not colar_tom_aco_inox__dn50a150xdnr20_estoque.empty:
                            self.tb_materiais.InsertRows(
                                str(ultima_linha_material))
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "ETAPA", sap_etapa_material
                            )
                            # Adiciona COLAR TOMADA ACO INOX DN50A150 X DNR20.
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "MATERIAL", "30004702"
                            )
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "QUANT", "1"
                            )
                            self.tb_materiais.setCurrentCell(
                                ultima_linha_material, "QUANT"
                            )
                            ultima_linha_material = ultima_linha_material + 1

    def receita_reparo_de_rede_de_agua(self):
        '''Padrão de materiais na classe CRA.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        # ABRACADEIRA FF REPARO TUBO DN100 LMIN=150
        abrac_ff_reparo_dn100_l150_estoque = self.estoque[
            self.estoque['Material'] == '30008103']
        # TUBO PBA DN 50 1,00 MPA JEI/JERI CM 6M
        tubo_dn50_estoque = self.estoque[self.estoque['Material']
                                         == '30028862']
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
                    if not abrac_ff_reparo_dn100_l150_estoque.empty:
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
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

            if tubo_pba_dn50 is False and self.diametro_rede == '50' \
                    and not tubo_dn50_estoque.empty:
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
                self.tb_materiais, self.contrato, self.estoque)

    def receita_troca_de_conexao_de_ligacao_de_agua(self):
        '''Padrão de materiais na classe Troca de Conexão de Ligação de Água.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        # CONEXOES MET LIGACOES FEMEA DN 20
        con_met_femea_dn20_estoque = self.estoque[self.estoque['Material'] == '30002394']
        # REGISTRO METALICO RAMAL PREDIAL DN 20
        reg_met_predial_dn20_estoque = self.estoque[self.estoque['Material'] == '30006747']
        if sap_material is None:
            ultima_linha_material = 0
            if not con_met_femea_dn20_estoque.empty:
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

            if not reg_met_predial_dn20_estoque.empty:
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

            if '30002394' not in material_lista \
                    and not con_met_femea_dn20_estoque.empty:
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

            if '30006747' not in material_lista \
                    and not reg_met_predial_dn20_estoque.empty:
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
                self.tb_materiais, self.contrato, self.estoque)

    def receita_reparo_de_ramal_de_agua(self):
        '''Padrão de materiais no reparo de Ligação de Água.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        tubo_pead_dn20_estoque = self.estoque[self.estoque['Material'] == '30001848']
        if sap_material is None and not tubo_pead_dn20_estoque.empty:
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

            if '30001848' not in material_lista and not tubo_pead_dn20_estoque.empty:
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
                self.tb_materiais, self.contrato, self.estoque)
