# hidrometro_material.py
'''Módulo dos materiais de família Hidrômetro.'''
from sap_connection import connect_to_sap
from wms import testa_material_sap
from wms import materiais_contratada
from wms import lacre_material


class HidrometroMaterial:
    '''Classe de materiais do hidrômetro.'''

    def __init__(self, int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais,
                 contrato,
                 estoque) -> None:
        self.int_num_lordem = int_num_lordem
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato
        self.estoque = estoque

    def receita_hidrometro(self):
        '''Padrão de materiais na classe Hidrômetro.'''
        session = connect_to_sap()
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        if self.hidro is None:
            return
        hidro_instalado = self.hidro
        if sap_material is None:
            if hidro_instalado is not None:
                print("Tem hidro, mas não foi vinculado!")
                ultima_linha_material = 0
                hidro_y = 'Y'
                hidro_a = 'A'
                # Hidrômetro atual.
                hidro_instalado = hidro_instalado.upper()
                # Mata-burro pra hidro.
                if hidro_instalado.startswith(hidro_y):
                    cod_hidro_instalado = '50000108'
                if hidro_instalado.startswith(hidro_a):
                    cod_hidro_instalado = '50000530'
                # Colocar lacre.
                hidro_estoque = self.estoque[self.estoque['Material']
                                             == cod_hidro_instalado]
                lacre_estoque = self.estoque[self.estoque['Material']
                                             == '50001070']
                if not hidro_estoque.empty and not lacre_estoque.empty:
                    print(f"Incluindo hidro: {cod_hidro_instalado} e lacre.")
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.operacao
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "50001070"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1
                    # Colocar hidrometro
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.operacao
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", cod_hidro_instalado
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1
        else:
            num_material_linhas = self.tb_materiais.RowCount
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas
            hidro_y = 'Y'
            hidro_a = 'A'
            sap_hidro = []
            # Hidrômetro atual.
            hidro_instalado = hidro_instalado.upper()

            # Mata-burro pra hidro.
            if hidro_instalado.startswith(hidro_y):
                cod_hidro_instalado = '50000108'
            if hidro_instalado.startswith(hidro_a):
                cod_hidro_instalado = '50000530'

            hidro_estoque = self.estoque[self.estoque['Material']
                                         == cod_hidro_instalado]
            # Variável para controlar se o hidrômetro já foi adicionado
            hidro_adicionado = False
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                if sap_material == '50000530' and '50000530' != cod_hidro_instalado:
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                else:
                    sap_hidro.append(sap_material)

            if cod_hidro_instalado in sap_hidro:
                self.tb_materiais.pressToolbarButton("&FIND")
                session.findById(
                    "wnd[1]/usr/txtGS_SEARCH-VALUE").text = cod_hidro_instalado
                session.findById(
                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[1]").sendVKey(12)
                quantidade = self.tb_materiais.GetCellValue(
                    self.tb_materiais.CurrentCellRow, "QUANT"
                )
                qtd_float = float(quantidade.replace(",", "."))

                if qtd_float >= 2.00 and not hidro_estoque.empty:
                    print("Quantidade de hidro errada, consertando.")
                    self.tb_materiais.modifyCell(
                        self.tb_materiais.CurrentCellRow, "QUANT", "1")
                    self.tb_materiais.setCurrentCell(
                        self.tb_materiais.CurrentCellRow, "QUANT")

                hidro_adicionado = True

            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(
                    n_material, "ETAPA")

                if sap_material == ('50000328', '50000263', '50000350',
                                    '50000064', '50000260', '50000261'):
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

                if self.hidro is not None:
                    if hidro_adicionado is True:
                        print("Hidro já adicionado.")
                    else:
                        if sap_material == cod_hidro_instalado:
                            print(
                                f"Hidro foi incluso corretamente: {cod_hidro_instalado}")
                            self.tb_materiais.pressToolbarButton("&FIND")
                            session.findById(
                                "wnd[1]/usr/txtGS_SEARCH-VALUE").text = cod_hidro_instalado
                            session.findById(
                                "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                            session.findById("wnd[1]").sendVKey(0)
                            session.findById("wnd[1]").sendVKey(12)
                            quantidade = self.tb_materiais.GetCellValue(
                                self.tb_materiais.CurrentCellRow, "QUANT"
                            )
                            qtd_float = float(quantidade.replace(",", "."))

                            if qtd_float >= 2.00 and not hidro_estoque.empty:
                                print("Quantidade de hidro errada, consertando.")
                                self.tb_materiais.modifyCell(
                                    self.tb_materiais.CurrentCellRow, "QUANT", "1")
                                self.tb_materiais.setCurrentCell(
                                    self.tb_materiais.CurrentCellRow, "QUANT")
                            # Hidrômetro foi adicionado
                            hidro_adicionado = True

                        elif sap_material == '50000108' \
                                and sap_material != cod_hidro_instalado and not hidro_estoque.empty:
                            print(
                                "Hidro inserido incorretamente."
                                + f"\nIncluindo o informado: {cod_hidro_instalado}")
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True)
                            self.tb_materiais.InsertRows(
                                str(ultima_linha_material))
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "ETAPA", sap_etapa_material)
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "QUANT", "1")
                            self.tb_materiais.setCurrentCell(
                                ultima_linha_material, "QUANT")
                            ultima_linha_material = ultima_linha_material + 1
                            hidro_adicionado = True  # Hidrômetro foi adicionado

                        elif sap_material == '50000530' \
                                and sap_material != cod_hidro_instalado and not hidro_estoque.empty:
                            print(
                                "Hidro inserido incorretamente."
                                + f"\nIncluindo o informado: {cod_hidro_instalado}")
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True)
                            self.tb_materiais.InsertRows(
                                str(ultima_linha_material))
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "ETAPA", sap_etapa_material)
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "QUANT", "1")
                            self.tb_materiais.setCurrentCell(
                                ultima_linha_material, "QUANT")
                            ultima_linha_material = ultima_linha_material + 1
                            hidro_adicionado = True  # Hidrômetro foi adicionado

                        elif sap_material == '50000350' \
                                and sap_material != cod_hidro_instalado and not hidro_estoque.empty:
                            print(
                                "Hidro inserido incorretamente."
                                + f"\nIncluindo o informado: {cod_hidro_instalado}")
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True)
                            self.tb_materiais.InsertRows(
                                str(ultima_linha_material))
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "ETAPA", sap_etapa_material)
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                            self.tb_materiais.modifyCell(
                                ultima_linha_material, "QUANT", "1")
                            self.tb_materiais.setCurrentCell(
                                ultima_linha_material, "QUANT")
                            ultima_linha_material = ultima_linha_material + 1
                            hidro_adicionado = True  # Hidrômetro foi adicionado

            if self.hidro is not None \
                    and hidro_adicionado is False and not hidro_estoque.empty:
                print(
                    "Não foi inserido hidro, "
                    + f"incluindo o informado: {cod_hidro_instalado}")
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", sap_etapa_material)
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1")
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT")
                ultima_linha_material = ultima_linha_material + 1
                hidro_adicionado = True  # Hidrômetro foi adicionado

            # Materiais do Global.
            materiais_contratada.materiais_contratada(
                self.tb_materiais, self.contrato, self.estoque)
            lacre_material.caca_lacre(
                self.tb_materiais, self.operacao, self.estoque)

    def receita_desinclinado_hidrometro(self):
        '''Padrão de materiais na classe Hidrômetro Desinclinado.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        lacre_estoque = self.estoque[self.estoque['Material']
                                     == '50001070']
        if sap_material is None and not lacre_estoque.empty:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.operacao
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "50001070"
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1

        # Materiais do Global.
        materiais_contratada.materiais_contratada(
            self.tb_materiais, self.contrato, self.estoque)
