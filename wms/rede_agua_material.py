# hidrometro_material.py
'''Módulo dos materiais de família Rede de Água.'''
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from wms import testa_material_sap
from wms import materiais_contratada


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

class RedeAguaMaterial:
    '''Classe de materiais de CRA.'''
    def __init__(self, int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 tb_materiais) -> None:
        self.int_num_lordem = int_num_lordem
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais

    def receita_reparo_de_rede_de_agua(self):
        '''Padrão de materiais na classe CRA.'''
        sap_material = testa_material_sap.testa_material_sap(self.int_num_lordem, self.tb_materiais)
        abracadeira_dn75 = False
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
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")

                if sap_material == '30005088':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
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
                # COTOVELO 90 GR PVC BB JE ESG PRED DN 150 é material de esgoto.
                if sap_material == '30011136':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

                if sap_material in ('30004097', '30002152', '30002151'):
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )

                if sap_material in ('30004097', '30002152', '30002151') \
                and abracadeira_dn75 is False:
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona ABRACADEIR FF REPARO TUBO DN100 LMIN=150.
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

            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais)
