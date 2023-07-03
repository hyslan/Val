# almoxarifado.py
'''Módulo dos materiais contratada e SABESP.'''
import sys
import pywintypes
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from wms import corte_restab_material
from wms import hidrometro_material
from wms import ramal_agua_material
from wms import rede_agua_material
from wms import rede_esgoto_material

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

class Almoxarifado:
    '''Área de todos materiais obrigatórios por TSE.'''
    def __init__(self,
                 int_num_lordem,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede) -> None:
        self.hidro = hidro
        self.int_num_lordem = int_num_lordem
        # 0 - tse, 1 - etapa tse, 2 - id match case
        self.identificador = identificador
        self.operacao = operacao
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede

    def aba_materiais(self):
        '''Função habilita aba de materiais no sap'''
        print("****Processo de Materiais****")
        session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM").select()
        tb_materiais = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9030/cntlCC_MATERIAIS/shellcont/shell")
        return tb_materiais

    def testa_material_sap(self, tb_materiais):
        '''Módulo de verificar materiais inclusos na ordem.'''
        try:
            sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
            print("Tem material vinculado.")
            return sap_material
        # pylint: disable=E1101
        except pywintypes.com_error:
            material_obs = planilha.cell(row=self.int_num_lordem, column=3)
            material_obs.value = "Sem Material Vinculado"
            print("Sem material vinculado.")
            lista.save('lista.xlsx')  # salva Planilha
            return None

    def inspecao(self, tb_materiais):
        '''Seleciona a Classe da TSE correta.'''
        sondagem = [
        '591000',
        '567000',
        '321000',
        '283000'
        ]
        if self.identificador[0] in sondagem:
            pass
        else:
            match self.identificador[2]:
                case "hidrometro":
                    material = hidrometro_material.HidrometroMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                        )
                    print("Aplicando a receita de hidrômetro.")
                    material.receita_hidrometro()

                case "cavalete":
                    sap_material = self.testa_material_sap(tb_materiais)
                    if sap_material is not None:
                        self.materiais_contratada(tb_materiais)
                    else:
                        return

                case "religacao":
                    material = corte_restab_material.CorteRestabMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de religação.")
                    material.receita_religacao()

                case "supressao":
                    material = corte_restab_material.CorteRestabMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de Supressão.")

                    material.receita_supressao()
                case "ramal_agua" | "tra" | "ligacao_agua":
                    material = ramal_agua_material.LigacaoAguaMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de Troca de Conexão "
                            + "de Ligação de Água.")
                    material.receita_troca_de_conexao_de_ligacao_de_agua()

                case "rede_agua":
                    material = rede_agua_material.RedeAguaMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de Reparo de Rede de Água")
                    material.receita_reparo_de_rede_de_agua()

                case "ligacao_esgoto":
                    material = rede_esgoto_material.RedeEsgotoMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de Reparo de Ramal de Esgoto")
                    material.receita_reparo_de_ramal_de_esgoto()

                case "rede_esgoto":
                    material = rede_esgoto_material.RedeEsgotoMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de Reparo de Rede de Esgoto")
                    material.receita_reparo_de_rede_de_esgoto()

                case "preservacao":
                    pass

                case _:
                    print("Classe não identificada.")
                    sys.exit()
        return

    def materiais_contratada(self, tb_materiais):
        '''Módulo de materiais da NOVASP.'''
        num_material_linhas = tb_materiais.RowCount  # Conta as Rows
        # Número da Row do Grid Materiais do SAP
        n_material = 0
        ultima_linha_material = num_material_linhas
        # Loop do Grid Materiais.
        for n_material in range(num_material_linhas):
            # Pega valor da célula 0
            sap_material = tb_materiais.GetCellValue(
                n_material, "MATERIAL")
            sap_etapa_material = tb_materiais.GetCellValue(
                n_material, "ETAPA")
            # Verifica se está na lista tb_contratada
            if sap_material in tb_contratada:
                # Marca Contratada
                tb_materiais.modifyCheckbox(
                    n_material, "CONTRATADA", True)
                print(f"Linha do material: {n_material}, "
                        + f"Material: {sap_material}")
                continue
            if sap_material in ('50000328', '50001070'):
                # Remove o lacre bege antigo.
                tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
                tb_materiais.InsertRows(str(ultima_linha_material))
                tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", sap_etapa_material
                )
                tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "50000263"
                )
                tb_materiais.modifyCell(
                    ultima_linha_material, "QUANT", "1"
                )
                tb_materiais.setCurrentCell(
                    ultima_linha_material, "QUANT"
                )
                ultima_linha_material = ultima_linha_material + 1

            try:
                if sap_material == '10014709':
                    # Marca Contratada
                    tb_materiais.modifyCheckbox(
                        n_material, "CONTRATADA", True)
                    print("Aslfato frio da NOVASP por enquanto.")
            # pylint: disable=E1101
            except pywintypes.com_error:
                print(f"Etapa: {sap_etapa_material} - Asfalto frio já foi retirado.")

            try:
                if sap_material == '30028856':
                    # Marca Contratada
                    tb_materiais.modifyCheckbox(
                        n_material, "CONTRATADA", True)
                    print("TUBO ESG DN 100 da NOVASP por enquanto.")
            # pylint: disable=E1101
            except pywintypes.com_error:
                print(f"Etapa: {sap_etapa_material} - TUBO ESG DN 100 já foi retirado.")

    def caca_lacre(self, tb_materiais):
        '''Módulo de procurar lacres no grid de materiais.'''
        num_material_linhas = tb_materiais.RowCount  # Conta as Rows
        # Número da Row do Grid Materiais do SAP
        n_material = 0
        procura_lacre = []
        ultima_linha_material = num_material_linhas
        # Loop do Grid Materiais.
        for n_material in range(num_material_linhas):
            # Pega valor da célula 0
            sap_material = tb_materiais.GetCellValue(
                n_material, "MATERIAL")
            procura_lacre.append(sap_material)
        if sap_material == '50000328' or '500001070' not in procura_lacre:
            # Remove o lacre e inclui o lacre novo, apenas se não tiver.
            tb_materiais.modifyCheckbox(
                n_material, "ELIMINADO", True
            )
            tb_materiais.InsertRows(str(ultima_linha_material))
            tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            tb_materiais.modifyCell(
                ultima_linha_material, "MATERIAL", "50000263"
            )
            tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1


def materiais(int_num_lordem,
              hidro_instalado,
              operacao,
              identificador,
              diametro_ramal,
              diametro_rede
              ):
    '''Função dos materiais de acordo com a TSE pai.'''
    servico = Almoxarifado(int_num_lordem,
                           hidro_instalado,
                           operacao,
                           identificador,
                           diametro_ramal,
                           diametro_rede
                           )
    tb_materiais = servico.aba_materiais()
    servico.inspecao(tb_materiais)
