'''Módulo dos materiais de família Poço.'''
from wms import testa_material_sap
from wms import materiais_contratada


class PocoMaterial:
    '''Classe de materiais de Poço.'''

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

    def receita_caixa_de_parada(self):
        '''Padrão de materiais no módulo caixa de parada.'''
        sap_material = testa_material_sap.testa_material_sap(
            self.int_num_lordem, self.tb_materiais)
        tampao_estoque = self.estoque[self.estoque['Material'] == '30003442']
        if sap_material is None and not tampao_estoque.empty:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.operacao
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "30003442"
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
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

                if sap_material == '30029526' \
                        and self.contrato == "4600041302":
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

            # Materiais do Global.
            materiais_contratada.materiais_contratada(
                self.tb_materiais, self.contrato, self.estoque)
