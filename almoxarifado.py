# almoxarifado.py
'''Módulo dos materiais contratada e SABESP.'''
import sys
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from wms import corte_restab_material
from wms import hidrometro_material
from wms import ramal_agua_material
from wms import rede_agua_material
from wms import rede_esgoto_material
from wms import cavalete_material
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

    def inspecao(self, tb_materiais):
        '''Seleciona a Classe da TSE correta.'''
        sondagem = [
            '591000',
            '567000',
            '321000',
            '283000'
        ]
        if self.identificador[0] in sondagem:
            materiais_contratada.materiais_contratada(tb_materiais)
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

                case "desinclinado":
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
                    material.receita_desinclinado_hidrometro()

                case "cavalete":
                    material = cavalete_material.CavaleteMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita de Cavalete.")
                    material.receita_cavalete()

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
                    print("Aplicando a receita "
                          + "de Ligação (Ramal) de Água.")
                    material.receita_troca_de_conexao_de_ligacao_de_agua()

                case "reparo_ramal_agua":
                    material = ramal_agua_material.LigacaoAguaMaterial(
                        self.int_num_lordem,
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais
                    )
                    print("Aplicando a receita - reparo ramal de água")
                    material.receita_reparo_de_ramal_de_agua()

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
                    print("Aplicando a receita de Rede de Água")
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
                    print("Aplicando a receita de Ramal de Esgoto")
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
                    print("Aplicando a receita de Rede de Esgoto")
                    material.receita_reparo_de_rede_de_esgoto()

                case "preservacao":
                    pass

                case _:
                    print("Classe não identificada.")
                    sys.exit()
        return


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
