# hidrometro_material.py
"""Módulo dos materiais de família Hidrômetro."""
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
        """Tipo de hidrômetro."""
        hidro_y = 'Y'
        hidro_a = 'A'
        hidro_b = 'B'
        hidro_f = 'F'
        self.hidro = self._hidro.upper()
        if self._hidro is not None:
            if self._hidro.startswith(hidro_y):
                return '50000108'
            if self._hidro.startswith(hidro_a):
                return '50000530'
            if self._hidro.startswith(hidro_b):
                return '50000387'
            if self._hidro.startswith(hidro_f):
                return '50000025'
        return None

    def _set_hidro(self, ultima_linha_material, etapa, cod_hidro_instalado):
        self.tb_materiais.modifyCheckbox(
            self.tb_materiais.CurrentCellRow, "ELIMINADO", True)
        self.tb_materiais.InsertRows(
            str(ultima_linha_material))
        self.tb_materiais.modifyCell(
            ultima_linha_material, "ETAPA", etapa[0])
        self.tb_materiais.modifyCell(
            ultima_linha_material, "MATERIAL", cod_hidro_instalado)
        self.tb_materiais.modifyCell(
            ultima_linha_material, "QUANT", "1")
        self.tb_materiais.setCurrentCell(
            ultima_linha_material, "QUANT")
        ultima_linha_material = ultima_linha_material + 1
        return ultima_linha_material

    def _search_hidro(self, cod_hidro_instalado, num_material_linhas, hidro_estoque):
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
            sap_hidro = np.append(sap_hidro, [sap_material, sap_etapa_material], axis=0)

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
                print(
                    "Hidro inserido incorretamente."
                    + f"\nIncluindo o informado: {cod_hidro_instalado}")
                etapa = sap_hidro[sap_hidro[:, 0] == '50000108', 1]
                btn_busca_material(self.tb_materiais,
                                   self.session, '50000108')
                ultima_linha_material = self._set_hidro(
                    ultima_linha_material, etapa[0], cod_hidro_instalado)

            if '50000530' in sap_hidro[:, 0] \
                    and '50000530' != cod_hidro_instalado and not hidro_estoque.empty:
                print(
                    "Hidro inserido incorretamente."
                    + f"\nIncluindo o informado: {cod_hidro_instalado}")
                etapa = sap_hidro[sap_hidro[:, 0] == '50000108', 1]
                btn_busca_material(self.tb_materiais,
                                   self.session, '50000530')
                ultima_linha_material = self._set_hidro(
                    ultima_linha_material, etapa[0], cod_hidro_instalado)

            if '50000350' in sap_hidro[:, 0] \
                    and '50000350' != cod_hidro_instalado and not hidro_estoque.empty:
                print(
                    "Hidro inserido incorretamente."
                    + f"\nIncluindo o informado: {cod_hidro_instalado}")
                etapa = sap_hidro[sap_hidro[:, 0] == '50000108', 1]
                btn_busca_material(self.tb_materiais,
                                   self.session, '50000350')
                ultima_linha_material = self._set_hidro(
                    ultima_linha_material, etapa[0], cod_hidro_instalado)

            return True, sap_hidro

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
        # Hidrômetro atual.
        cod_hidro_instalado = self._hidro_type()

        hidro_estoque = self.estoque[self.estoque['Material']
                                     == cod_hidro_instalado]
        # Variável para controlar se o hidrômetro já foi adicionado
        hidro_adicionado = False

        hidro_adicionado, sap_hidro = self._search_hidro(
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

            if self._hidro is not None:
                if hidro_adicionado is True:
                    print("Hidro já adicionado.")
                else:
                    if sap_material == '50000108' \
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
        if self._hidro is not None \
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
