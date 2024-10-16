# hidrometro_material.py
"""Módulo dos materiais de família Corte/Relig."""

from __future__ import annotations

import typing

from python.src.wms import lacre_material, materiais_contratada, testa_material_sap

if typing.TYPE_CHECKING:
    from pandas import DataFrame
    from win32com.client import CDispatch


class CorteRestabMaterial:
    """Classe de materiais de religação e supressão."""

    def __init__(
        self,
        hidro: str,
        operacao: str,
        identificador: tuple[str, str, str, list[str], list[str]],  # Unique Array key
        diametro_ramal: str,
        diametro_rede: str,
        tb_materiais: CDispatch,
        contrato: str,
        estoque: DataFrame,
        session: CDispatch,
    ) -> None:
        """Método de inicialização da classe CorteRestabMaterial.

        Args:
        ----
            hidro (str): Número de Série do Hidrometro.
            operacao (str): Etapa Pai
            identificador (tuple[str, str, str]): TSE, Etapa TSE, ID Match Case do inspector de Almoxarixado.py
            diametro_ramal (str): Tamanho do diâmetro do ramal.
            diametro_rede (str): Tamanho do diâmetro da rede.
            tb_materiais (CDispatch): GRID de Materiais.
            contrato (str): Número do contrato.
            estoque (DataFrame): Tabela de Estoque.
            session (CDispatch): Sessão do SAP.

        """
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato
        self.estoque = estoque
        self.session = session
        self.list_contratada = materiais_contratada.lista_materiais()

    def receita_religacao(self) -> None:
        """Padrão de materiais na classe Religação."""
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)
        lacre_estoque = self.estoque[self.estoque["Material"] == "50001070"]
        if sap_material is None and not lacre_estoque.empty:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material,
                "ETAPA",
                self.operacao,
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material,
                "MATERIAL",
                "50001070",
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material,
                "QUANT",
                "1",
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material,
                "QUANT",
            )
            ultima_linha_material = ultima_linha_material + 1
        else:
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
                material_estoque = self.estoque[self.estoque["Material"] == sap_material]

                # Retirar hidro vinculado em religação.
                if sap_material in ("50000108", "50000530"):
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if sap_material == "30029526" and self.contrato == "4600041302":
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if sap_material in ("30001848", "30007034"):
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if sap_material != "50001070" and sap_material not in self.list_contratada:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if material_estoque.empty:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )

            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais, self.contrato, self.estoque, self.session)
            # Caça lacre
            lacre_material.caca_lacre(self.tb_materiais, self.operacao, self.estoque, self.session)

    def receita_supressao(self) -> None:
        """Padrão de materiais para supressão."""
        materiais_receita = [
            "30029526",
            "10014709",
            "30003832",
            "50001070",
        ]
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)
        if sap_material is not None:
            material_lista = []
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
                material_lista.append(sap_material)
                material_estoque = self.estoque[self.estoque["Material"] == sap_material]

                if sap_material == "30029526" and self.contrato == "4600041302":
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if sap_material in ("30001848", "30007034"):
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if sap_material not in materiais_receita and sap_material not in self.list_contratada:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if material_estoque.empty:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )

            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais, self.contrato, self.estoque, self.session)
            # Caça lacre
            lacre_material.caca_lacre(self.tb_materiais, self.operacao, self.estoque, self.session)
