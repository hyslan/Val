"""Módulo dos materiais de família Poço."""


from python.src.wms import localiza_material, materiais_contratada, testa_material_sap


class PocoMaterial:
    """Classe de materiais de Poço."""

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

        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato
        self.estoque = estoque
        self.session = session

    def receita_caixa_de_parada(self) -> None:
        """Padrão de materiais no módulo caixa de parada e descoberto."""
        sap_material = testa_material_sap.testa_material_sap(
            self.tb_materiais)
        tampao_estoque = self.estoque[self.estoque["Material"] == "30003442"]
        # In case of 'Troca caixa de parada'
        if sap_material is None and not tampao_estoque.empty \
                and self.identificador[0] == "322000":
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.operacao,
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "30003442",
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1",
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT",
            )
            ultima_linha_material = ultima_linha_material + 1
        else:
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")

                if sap_material == "30029526" \
                        and self.contrato == "4600041302":
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True,
                    )

            # Materiais do Global.
            materiais_contratada.materiais_contratada(
                self.tb_materiais, self.contrato,
                self.estoque, self.session)

    def niv_pv_pi(self) -> None:
        """Material do bloco de nivelamento de PV e PI."""
        sap_material = testa_material_sap.testa_material_sap(
            self.tb_materiais)
        tampao_estoque = self.estoque[self.estoque["Material"] == "30032220"]
        if sap_material is not None and not tampao_estoque.empty:
            procura_tampao = []
            for n_material in range(self.tb_materiais.RowCount):
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                procura_tampao.append(sap_material)

            if "30032220" in procura_tampao:
                localiza_material.btn_busca_material(
                    self.tb_materiais, self.session, "30032220")
                qtd = self.tb_materiais.GetCellValue(
                    self.tb_materiais.CurrentCellRow, "QUANT",
                )
                qtd_float = float(qtd.replace(",", "."))
                if qtd_float >= 2.00 or qtd_float == 0.0 and not tampao_estoque.empty:
                    self.tb_materiais.modifyCell(
                        self.tb_materiais.CurrentCellRow, "QUANT", "1",
                    )
                    self.tb_materiais.setCurrentCell(
                        self.tb_materiais.CurrentCellRow, "QUANT",
                    )
                # Caso sem estoque
                if tampao_estoque.empty:
                    self.tb_materiais.modifyCheckbox(
                        self.tb_materiais.CurrentCellRow, "ELIMINADO", True,
                    )
