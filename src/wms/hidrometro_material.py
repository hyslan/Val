# hidrometro_material.py
"""Módulo dos materiais de família de serviços relacionados a Hidrômetro.
    Materiais obrigatórios:
        * HIDRÔMETRO;
        * LACRE;
    @AUTHOR= HYSLAN SILVA CRUZ
    @YEAR= 2023
"""
import numpy as np
from src.wms import testa_material_sap
from src.wms import materiais_contratada
from src.wms import lacre_material
from src.wms.localiza_material import btn_busca_material
from src import sql_view


class HidrometroMaterial:
    """Classe de materiais do hidrômetro."""

    def __init__(self,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais,
                 contrato,
                 estoque,
                 session) -> None:
        self._hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato
        self.estoque = estoque
        self.session = session
        self.list_contratada = materiais_contratada.lista_materiais()

    @property
    def hidro(self):
        return self._hidro

    @hidro.setter
    def hidro(self, hd):
        if isinstance(hd, str):
            self._hidro = hd
        else:
            raise ValueError("Wrong type, need to be string.")

    def _hidro_type(self):
        """Tipo de hidrômetro e seu fabricante de acordo com a matrícula
        em sua primeira leitura no campo hidrômetro instalado.
        Fonte de Referência:
        http://10.66.9.42/catalogomateriais/index.php/arquivos/pesquisa
        """
        self.hidro = self._hidro.upper()
        if self._hidro is not None:
            if self._hidro.startswith('Y'):
                # HIDROMETRO QN 0,75 M3/H (QMAX) CL B EQUIPADO SAIDA M-BUS
                # FABRICANTE ITRON -> ANO LOTE: 2016
                if self._hidro[3] == 'T':
                    return '50000247'
                return '50000108'
            if self._hidro.startswith('A'):
                # HIDROMETRO ULTRASSONICO DN20-QN 1,5 MBUS OU PULSO
                # FABRICANTE: HYDRUS
                if self._hidro[3] == 'B':
                    return '50000112'
                # Hidro ultrassônico SAGA
                if self._hidro[3] == 'G':
                    return '50000387'
                # HIDROMETRO TAQUIMETRICO DN 20 - CLASSE B -QN 1.5 M3/H -> ANO LOTE: 2017
                # HIDRÔMETRO VOLUMÉTRICO DN 20 - QN 1 5 ANTI SUPER IMA - VOLUMETRICO -> ANO LOTE: 2017
                # HIDRÔMETRO DN20 - QN 1,5 - VOLUM - ANTI S.I- R.INCLINADO -> ANO LOTE: ATUAL
                # FABRICANTE: LAO
                if self._hidro[3] == 'L':
                    # ANTIGO COD. SAP: '50000321'
                    # ANTIGO COD. SAP: '50000326'
                    return '50000530'
                # HIDROMETRO TAQUIMETRICO DN 20 - QN 1,5M3/H - CL C INCL
                # FABRICANTE: ELSTER
                if self._hidro[3] == 'N':
                    return '50000230'
                # HIDROMETRO VOLUMETRICO QN 1,5 M3/H CL C PRE-EQUIPADO
                # FABRICANTE: AQUADIS+
                if self._hidro[3] == 'S':
                    return '50000140'
                return '50000530'
            if self._hidro.startswith('B'):
                return '50000387'
            if self._hidro.startswith('D'):
                # Hidro Fabricante ZENER
                if self._hidro[3] == 'Z':
                    return '50000221'
            if self._hidro.startswith('F'):
                return '50000025'
        return None

    def _delete_row(self):
        """This line of code is used to modify a checkbox in a table.
        The 'modifyCheckbox' method is called on the 'tb_materiais' object.
        The method takes three arguments:
        1. The row of the cell, which is the current cell row in this case.
        2. The name of the column, which is "ELIMINADO" in this case.
        3. The new value for the checkbox, which is set to True in this case.
        This effectively marks the current row as "eliminated" in the 'tb_materiais' table.
        """
        self.tb_materiais.modifyCheckbox(
            self.tb_materiais.CurrentCellRow, "ELIMINADO", True)

    def _set_hidro(self, ultima_linha_material, etapa, cod_hidro_instalado):
        """Set the correct hydrometer"""
        self.tb_materiais.SetCurrentCell(0, 'ETAPA')
        self.tb_materiais.InsertRows(
            str(ultima_linha_material))
        self.tb_materiais.modifyCell(
            ultima_linha_material, "ETAPA", etapa) # Etapa do serviço principal Hidrometro
        self.tb_materiais.modifyCell(
            ultima_linha_material, "MATERIAL", cod_hidro_instalado)
        self.tb_materiais.modifyCell(
            ultima_linha_material, "QUANT", "1")
        self.tb_materiais.setCurrentCell(
            ultima_linha_material, "QUANT")
        ultima_linha_material = ultima_linha_material + 1
        return ultima_linha_material

    def _search_hidro(self, cod_hidro_instalado, num_material_linhas, hidro_estoque):
        """Search for the hydrometer in the SAP table.
        If the hydrometer is not found, it is added to the table.
        """
        ultima_linha_material = num_material_linhas
        sap_hidro = np.empty((0, 2), dtype=str)
        for n_material in range(num_material_linhas):
            sap_material = self.tb_materiais.GetCellValue(
                n_material, "MATERIAL")
            sap_etapa_material = self.tb_materiais.GetCellValue(
                n_material, "ETAPA")
            if sap_material == '50000530' and '50000530' != cod_hidro_instalado:
                self.tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
            sap_hidro = np.append(sap_hidro,
                                  np.array([[sap_material, sap_etapa_material]]),
                                  axis=0)

        if cod_hidro_instalado in sap_hidro[:, 0]:
            btn_busca_material(self.tb_materiais,
                               self.session, cod_hidro_instalado)
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

        if '50000108' in sap_hidro[:, 0] \
                and '50000108' != cod_hidro_instalado and not hidro_estoque.empty:
            btn_busca_material(self.tb_materiais,
                               self.session, '50000108')
            self._delete_row()

        if '50000530' in sap_hidro[:, 0] \
                and '50000530' != cod_hidro_instalado and not hidro_estoque.empty:

            btn_busca_material(self.tb_materiais,
                               self.session, '50000530')
            self._delete_row()

        if '50000350' in sap_hidro[:, 0] \
                and '50000350' != cod_hidro_instalado and not hidro_estoque.empty:
            btn_busca_material(self.tb_materiais,
                               self.session, '50000350')
            self._delete_row()

        if not hidro_estoque.empty \
                and cod_hidro_instalado not in sap_hidro[:, 0]:
            ultima_linha_material = self._set_hidro(
                ultima_linha_material, self.operacao, cod_hidro_instalado)

        return True, ultima_linha_material

    def receita_hidrometro(self):
        """Padrão de materiais na classe Hidrômetro."""
        usr = self.session.findById("wnd[0]/usr")
        ordem = usr.findById("txtGS_HEADER-NUM_ORDEM").text
        lacre_estoque = self.estoque[self.estoque['Material']
                                     == '50001070']
        sap_material = testa_material_sap.testa_material_sap(
            self.tb_materiais)
        num_material_linhas = self.tb_materiais.RowCount
        ultima_linha_material = 0
        if self._hidro is None:
            print("Hidro não informado. Buscando no Geocall.")
            sql = sql_view.Tabela(ordem, "")
            self.hidro = sql.get_new_hidro()
            if self._hidro is None:
                print("Hidro não encontrado.")
                return
        cod_hidro_instalado = self._hidro_type()
        hidro_estoque = self.estoque[self.estoque['Material']
                                     == cod_hidro_instalado]
        print(f"Cod hidro: {cod_hidro_instalado}")

        if sap_material is None:
            print("Tem hidro, mas não foi vinculado!")
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

        # Número da Row do Grid Materiais do SAP
        ultima_linha_material = num_material_linhas
        # Variável para controlar se o hidrômetro já foi adicionado
        hidro_adicionado = False

        hidro_adicionado, ultima_linha_material = self._search_hidro(
            cod_hidro_instalado, num_material_linhas, hidro_estoque)

        # Loop do Grid Materiais.
        for n_material in range(num_material_linhas):
            sap_material = self.tb_materiais.GetCellValue(
                n_material, "MATERIAL")
            sap_etapa_material = self.tb_materiais.GetCellValue(
                n_material, "ETAPA")
            material_estoque = self.estoque[self.estoque['Material'] == sap_material]

            if sap_material == ('50000328', '50000263', '50000350',
                                '50000064', '50000260', '50000261',
                                '30001848', '30007034'):
                self.tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )

            # Only Hidro and Lacre.
            if sap_material not in (cod_hidro_instalado, '50001070') \
                    and sap_material not in self.list_contratada:
                self.tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
            # Material not in stock.
            if material_estoque.empty:
                self.tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )

        # Materiais do Global.
        materiais_contratada.materiais_contratada(
            self.tb_materiais, self.contrato,
            self.estoque, self.session)
        lacre_material.caca_lacre(
            self.tb_materiais, self.operacao,
            self.estoque, self.session)

    def receita_desinclinado_hidrometro(self):
        """Padrão de materiais na classe Hidrômetro Desinclinado."""
        sap_material = testa_material_sap.testa_material_sap(
            self.tb_materiais)
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

        else:
            num_material_linhas = self.tb_materiais.RowCount
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(
                    n_material, "ETAPA")
                material_estoque = self.estoque[self.estoque['Material'] == sap_material]

                if not sap_material == '50001070' \
                        and sap_material not in self.list_contratada:
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                # Material not in stock.
                if material_estoque.empty:
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

        # Materiais do Global.
        materiais_contratada.materiais_contratada(
            self.tb_materiais, self.contrato,
            self.estoque, self.session)
        # Caça lacre
        lacre_material.caca_lacre(
            self.tb_materiais, self.operacao,
            self.estoque, self.session)
