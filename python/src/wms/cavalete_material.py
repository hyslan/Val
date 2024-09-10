# cavalete_material.py
"""Módulo dos materiais de família Cavalete."""

from __future__ import annotations

import typing

from rich.console import Console

from python.src.wms import lacre_material, materiais_contratada, testa_material_sap

if typing.TYPE_CHECKING:
    import win32com.client as win32
    from pandas import DataFrame

console = Console()


class CavaleteMaterial:
    """Classe de materiais de cavalete."""

    def __init__(
        self,
        hidro: str,
        operacao: str,
        identificador: tuple[str, str, str, list[str], list[str]],  # Unique Array key
        diametro_ramal: str,
        diametro_rede: str,
        tb_materiais: win32.CDispatch,
        contrato: str,
        estoque: DataFrame,
        session: win32.CDispatch,
    ) -> None:
        """Método de inicialização da classe CavaleteMaterial.

        Args:
        ----
            hidro (str): Número de Série do Hidrometro.
            operacao (str): Etapa pai
            identificador (tuple[str, str, str]): TSE, Etapa TSE, ID Match Case do inspector de Almoxarixado.py
            diametro_ramal (str): Tamanho do diâmetro do ramal.
            diametro_rede (str): Tamanho do diâmetro da rede.
            tb_materiais (win32.CDispatch): GRID de Materiais.
            contrato (str): Número do contrato.
            estoque (DataFrame): Estoque da contratada.
            session (win32.CDispatch): Sessão do SAPGUI

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

    def receita_cavalete(self) -> None:
        """Padrão de materiais na classe Religação.

        Segue a risca os materiais da lista materiais_receita e os da contratada,
        exclui os materiais sem estoque.
        """
        list_contratada = materiais_contratada.lista_materiais()
        materiais_receita = [
            "30002394",
            "50000021",
            "50000178",
            "50001070",
            "30001346",
            "30004643",
            "30006747",
            "50000159",
            "50000472",
            "30001848",
            "30007896",
        ]
        materiais_lancados = []
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)
        lacre_estoque = self.estoque[self.estoque["Material"] == "50001070"]
        # Gambiarra
        if sap_material is None and lacre_estoque.empty:
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
                "0",
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material,
                "QUANT",
            )
            self.tb_materiais.modifyCheckBox(
                self.tb_materiais.CurrentCellRow,
                "ELIMINADO",
                True,
            )
            ultima_linha_material = ultima_linha_material + 1

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
            console.print(f"\nEtapa: {self.operacao}")
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(n_material, "ETAPA")
                materiais_lancados.append({"Material": sap_material, "Etapa": sap_etapa_material})
                material_estoque = self.estoque[self.estoque["Material"] == sap_material]

                if sap_material not in list_contratada:
                    console.print(f"\n{material_estoque}", style="italic green")

                if sap_material == "30029526" and self.contrato == "4600041302":
                    pass
                if (
                    sap_material not in materiais_receita
                    and sap_material not in list_contratada
                    and sap_etapa_material == self.operacao
                ):
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if material_estoque.empty and sap_material not in list_contratada:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )

            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais, self.contrato, self.estoque, self.session)
            lacre_material.caca_lacre(self.tb_materiais, self.operacao, self.estoque, self.session)
