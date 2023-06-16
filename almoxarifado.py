# almoxarifado.py
'''Módulo dos materiais contratada e SABESP.'''
import sys
import pywintypes
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets

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
    def __init__(self, int_num_lordem, hidro, operacao, identificador) -> None:
        self.hidro = hidro
        self.int_num_lordem = int_num_lordem
        self.identificador = identificador
        self.operacao = operacao

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
        match self.identificador:
            case "hidrometro":
                material = HidrometroMaterial(
                    self.int_num_lordem,
                    self.hidro,
                    self.operacao,
                    self.identificador,
                    tb_materiais
                    )
                print("Aplicando a receita de hidrômetro.")
                material.receita_hidrometro()
            case "cavalete":
                print("em construção.")
            case "religacao":
                return None
            case "supressao":
                return None
            case _:
                print("Classe não identificada.")
                sys.exit()
        return

    def materiais_contratada(self, tb_materiais):
        '''Módulo de materiais da NOVASP.'''
        num_material_linhas = tb_materiais.RowCount  # Conta as Rows
        print(f"Qtd de linhas de materiais: {num_material_linhas}")
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
            if sap_material == '50000328':
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

class HidrometroMaterial(Almoxarifado):
    '''Classe de materiais do hidrômetro.'''
    def __init__(self, int_num_lordem, hidro, operacao, identificador, tb_materiais):
        super().__init__(int_num_lordem, hidro, operacao, identificador)
        self.tb_materiais = tb_materiais
    def receita_hidrometro(self):
        '''Padrão de materiais na classe Hidrômetro.'''
        sap_material = super().testa_material_sap
        if sap_material is None:
            if hidro_instalado is not None:
                print("Tem hidro, mas não foi vinculado!")
                ultima_linha_material = 0
                hidro_y = 'Y'
                # Hidrômetro atual.
                hidro_instalado = hidro_instalado.upper()
                # Mata-burro pra hidro.
                if hidro_instalado.startswith(hidro_y):
                    cod_hidro_instalado = '50000108'
                else:
                    cod_hidro_instalado = '50000530'
                # Colocar lacre.
                print(f"Incluindo hidro: {cod_hidro_instalado} e lacre.")
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
        else:
            num_material_linhas = self.tb_materiais.RowCount
            print(f"Qtd de linhas de materiais: {num_material_linhas}")
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas
            hidro_y = 'Y'
            # Hidrômetro atual.
            self.hidro = self.hidro.upper()
            # Mata-burro pra hidro.
            if self.hidro.startswith(hidro_y):
                cod_hidro_instalado = '50000108'
            else:
                cod_hidro_instalado = '50000530'
            # Variável para controlar se o hidrômetro já foi adicionado
            hidro_adicionado = False
            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(
                    n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(
                    n_material, "ETAPA")
                # Verifica se está na lista tb_contratada
                if sap_material in tb_contratada:
                    # Marca Contratada
                    self.tb_materiais.modifyCheckbox(
                        n_material, "CONTRATADA", True)
                    print(f"Linha do material: {n_material}, "
                            + f"Material: {sap_material}")
                    continue
                if sap_material == '50000328':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", sap_etapa_material
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

                if  self.hidro is not None:
                    if hidro_adicionado is True:
                        print("Hidro já adicionado.")
                    else:
                        if sap_material == cod_hidro_instalado:
                            print(
                                f"Hidro foi incluso corretamente: {cod_hidro_instalado}")
                            # Hidrômetro foi adicionado
                            hidro_adicionado = True

                        elif sap_material == '50000108' and sap_material != cod_hidro_instalado:
                            print(
                                "Hidro inserido incorretamente."
                                + f"\nIncluindo o informado: {cod_hidro_instalado}")
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True)
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

                        elif sap_material == '50000530' and sap_material != cod_hidro_instalado:
                            print(
                                "Hidro inserido incorretamente."
                                + f"\nIncluindo o informado: {cod_hidro_instalado}")
                            self.tb_materiais.modifyCheckbox(
                                n_material, "ELIMINADO", True)
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

            if self.hidro is not None and hidro_adicionado is False:
                print(
                    f"Não foi inserido hidro, incluindo o informado: {cod_hidro_instalado}")
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", sap_etapa_material)
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                self.tb_materiais.modifyCell(ultima_linha_material, "QUANT", "1")
                self.tb_materiais.setCurrentCell(ultima_linha_material, "QUANT")
                ultima_linha_material = ultima_linha_material + 1
                hidro_adicionado = True  # Hidrômetro foi adicionado

def materiais(int_num_lordem, hidro_instalado, operacao, identificador):
    '''Função dos materiais de acordo com a TSE pai.'''
    servico = Almoxarifado(int_num_lordem, hidro_instalado, operacao, identificador)
    tb_materiais = servico.aba_materiais()
    servico.inspecao(tb_materiais)
    sap_material = servico.testa_material_sap(tb_materiais)
    if sap_material is not None:
        servico.materiais_contratada(tb_materiais)
        