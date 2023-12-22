# hidrometro_material.py
'''Módulo dos materiais de família Rede de Esgoto.'''
from wms import testa_material_sap
from wms import materiais_contratada
from wms import localiza_material
from sap_connection import connect_to_sap


class RedeEsgotoMaterial:
    '''Classe de materiais de CRE.'''

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
                 posicao_rede
                 ) -> None:
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

    def ramal_luva_correr(self):
        '''Saber diâmetro do ramal para luva correr'''
        match self.diametro_ramal:
            case 'DN_100':
                luva_correr = "30002797"
            case 'DN_150':
                luva_correr = "30001548"
            case 'DN_200':
                luva_correr = "30002798"
            case 'DN_300':
                luva_correr = "30008427"
            case 'DN_400':
                luva_correr = "30005898"
            case _:
                print("Diâmetro do ramal não informado.")
                luva_correr = None
                return luva_correr

        return luva_correr

    def curva(self):
        '''Curva 90 e 45'''
        match self.diametro_ramal:
            case 'DN_100':
                curva90 = "30002722"
                curva45 = "30005282"
            case 'DN_200':
                curva45 = "30005284"
                curva90 = "30005147"
            case 'DN_300':
                curva45 = "30005285"
                curva90 = "30005148"
            case _:
                print("Diâmetro do ramal não informado, retornando.")
                curva45 = None
                curva90 = None
                return curva45, curva90

        return curva45, curva90

    def ramal_junta(self):
        '''Saber diâmetro do ramal para junta'''
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
                print("Diâmetro do ramal não informado.")
                junta_esgoto = None
                junta_esgoto_adap = None
                return junta_esgoto, junta_esgoto_adap

        return junta_esgoto, junta_esgoto_adap

    def rede_junta(self):
        '''Saber diâmetro da rede para junta'''
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
                print("Diâmetro da rede não informado.")
                junta_esgoto = None
                junta_esgoto_adap = None
                return junta_esgoto, junta_esgoto_adap

        return junta_esgoto, junta_esgoto_adap

    def materiais_vigentes(self):
        '''Materiais com estoque'''
        session = connect_to_sap()
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        if sap_material is not None:
            luva_correr = self.ramal_luva_correr()
            curva45, curva90 = self.curva()
            num_material_linhas = self.tb_materiais.RowCount
            n_material = 0
            ultima_linha_material = num_material_linhas

            for material in self.df_materiais['Material']:
                match material:
                    case [curva90]:
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, curva90)
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case [curva45]:
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, curva45)
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    case [luva_correr]:
                        resultado = localiza_material.qtd_max(
                            material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(
                                self.tb_materiais, session, luva_correr)
                            localiza_material.qtd_correta(
                                self.tb_materiais, "1")
                    # TUBO PVC RIG JEI/JERI ESG DN 100 CM 6M
                    case '30028856':
                        codigo = '30028856'
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
                    case '30022469':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
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

                    case '30000139':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
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
                    case '30002858':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

                    case '30005616':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
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

                    case '30005894':
                        pass
                        # self.tb_materiais.modifyCheckbox(
                        #     n_material, "ELIMINADO", True
                        # )
                        # self.tb_materiais.InsertRows(str(ultima_linha_material))
                        # self.tb_materiais.modifyCell(
                        #     ultima_linha_material, "ETAPA", self.identificador[1]
                        # )
                        # # Adiciona LUVA CORRER BB ESG DN 100 vigente.
                        # self.tb_materiais.modifyCell(
                        #     ultima_linha_material, "MATERIAL", "30002797"
                        # )
                        # self.tb_materiais.modifyCell(
                        #     ultima_linha_material, "QUANT", "1"
                        # )
                        # self.tb_materiais.setCurrentCell(
                        #     ultima_linha_material, "QUANT"
                        # )
                        # ultima_linha_material = ultima_linha_material + 1

                    case '30005283':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
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

                    case '30007917':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", self.identificador[1]
                        )
                        # Adiciona o TUBO PVC Esgoto DN 100.
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "MATERIAL", "30028856"
                        )
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "QUANT", "1"
                        )
                        self.tb_materiais.setCurrentCell(
                            ultima_linha_material, "QUANT"
                        )
                        ultima_linha_material = ultima_linha_material + 1

                    case '30005329':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", self.identificador[1]
                        )
                        # Adiciona o Curva 90 GR PVC Esgoto DN 100.
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

                    case '30001657':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", self.identificador[1]
                        )
                        # Adiciona o SELIM AJUST TUBO PVC DN 150x100.
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "MATERIAL", "30007132"
                        )
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "QUANT", "1"
                        )
                        self.tb_materiais.setCurrentCell(
                            ultima_linha_material, "QUANT"
                        )
                        ultima_linha_material = ultima_linha_material + 1

                    case '30003816':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", self.identificador[1]
                        )
                        # Adiciona o TUBO PVC Esgoto DN 100.
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "MATERIAL", "30028856"
                        )
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "QUANT", "1"
                        )
                        self.tb_materiais.setCurrentCell(
                            ultima_linha_material, "QUANT"
                        )
                        ultima_linha_material = ultima_linha_material + 1

                    case '30005145':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", self.identificador[1]
                        )
                        # Adiciona o CURVA 90 GR PVC PB JE/JERI ESG DN 100.
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

                    case '30008877':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )
                        self.tb_materiais.InsertRows(
                            str(ultima_linha_material))
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "ETAPA", self.identificador[1]
                        )
                        # Adiciona o SELIM AJUST TUBO PVC E CER DN 150X100
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "MATERIAL", "30007132"
                        )
                        self.tb_materiais.modifyCell(
                            ultima_linha_material, "QUANT", "1"
                        )
                        self.tb_materiais.setCurrentCell(
                            ultima_linha_material, "QUANT"
                        )
                        ultima_linha_material = ultima_linha_material + 1

                    # Removendo JUNTA FLEX ESGOTO ADAP DN 100MMX100MM porque acabou.
                    case '30005615':
                        self.tb_materiais.modifyCheckbox(
                            n_material, "ELIMINADO", True
                        )

    def receita_reparo_de_rede_de_esgoto(self):
        '''Padrão de materiais na classe CRE.'''
        self.materiais_vigentes()
        # Materiais do Global.
        materiais_contratada.materiais_contratada(
            self.tb_materiais, self.contrato, self.estoque)

    def receita_reparo_de_ramal_de_esgoto(self):
        '''Padrão de materiais na classe Ramal de Esgoto.'''
        self.materiais_vigentes()
        # Materiais do Global.
        materiais_contratada.materiais_contratada(
            self.tb_materiais, self.contrato, self.estoque)

    def png(self):
        '''Método para PNG Esgoto'''
        self.materiais_vigentes()
        # Materiais do Global.
        materiais_contratada.materiais_contratada(
            self.tb_materiais, self.contrato, self.estoque)
