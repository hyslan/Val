# hidrometro_material.py
'''Módulo dos materiais de família Corte/Relig.'''
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from almoxarifado import Almoxarifado


session = connect_to_sap()
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
    _,
    *_,
) = load_worksheets()


class CorteRestabMaterial(Almoxarifado):
    '''Classe de materiais de religação e supressão.'''

    def __init__(self, int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais):
        super().__init__(int_num_lordem,
                         hidro,
                         operacao,
                         identificador,
                         diametro_ramal,
                         diametro_rede)
        self.tb_materiais = tb_materiais

    def receita_religacao(self):
        '''Padrão de materiais na classe Religação.'''
        sap_material = super().testa_material_sap(self.tb_materiais)
        if sap_material is None:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.operacao
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "50000263"
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
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                # Retirar hidro vinculado em religação.
                if sap_material in ('50000108', '50000530'):
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
            # Materiais do Global.
            self.materiais_contratada(self.tb_materiais)

    def receita_supressao(self):
        '''Padrão de materiais para supressão.'''
        sap_material = super().testa_material_sap(self.tb_materiais)
        if sap_material is not None:
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
                if sap_material == '30029526':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona o UNIAO P/TUBO PEAD DE 20 MM.
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30001865"
                    )
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "QUANT", "1"
                    )
                    self.tb_materiais.setCurrentCell(
                        ultima_linha_material, "QUANT"
                    )
                    ultima_linha_material = ultima_linha_material + 1
            # Materiais do Global.
            self.materiais_contratada(self.tb_materiais)
