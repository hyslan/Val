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
        # 0 - tse, 1 - etapa tse, 2 - id match case
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
        match self.identificador[2]:
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
                sap_material = self.testa_material_sap(tb_materiais)
                if sap_material is not None:
                    self.materiais_contratada(tb_materiais)
                else:
                    return

            case "religacao":
                material = CorteRestabMaterial(
                    self.int_num_lordem,
                    self.hidro,
                    self.operacao,
                    self.identificador,
                    tb_materiais
                )
                print("Aplicando a receita de religação.")
                material.receita_religacao()

            case "supressao":
                material = CorteRestabMaterial(
                    self.int_num_lordem,
                    self.hidro,
                    self.operacao,
                    self.identificador,
                    tb_materiais
                )
                print("Aplicando a receita de Supressão.")

                material.receita_supressao()
            case "ramal_agua" | "tra" | "ligacao_agua":
                material = LigacaoAguaMaterial(
                    self.int_num_lordem,
                    self.hidro,
                    self.operacao,
                    self.identificador,
                    tb_materiais
                )
                print("Aplicando a receita de Troca de Conexão "
                        + "de Ligação de Água.")
                material.receita_troca_de_conexao_de_ligacao_de_agua()

            case "rede_agua":
                material = RedeAguaMaterial(
                    self.int_num_lordem,
                    self.hidro,
                    self.operacao,
                    self.identificador,
                    tb_materiais
                )
                print("Aplicando a receita de Reparo de Rede de Água")
                material.receita_reparo_de_rede_de_agua()

            case "preservacao":
                return

            case _:
                print("Classe não identificada.")
                sys.exit()
        return

    def materiais_contratada(self, tb_materiais):
        '''Módulo de materiais da NOVASP.'''
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
        n_material = 0
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
                # Remove o lacre bege antigo.
                tb_materiais.modifyCheckbox(
                    n_material, "ELIMINADO", True
                )
            try:
                if sap_material == '10014709':
                    # Marca Contratada
                    tb_materiais.modifyCheckbox(
                        n_material, "CONTRATADA", True)
                    print("Aslfato frio da NOVASP por enquanto.")
            # pylint: disable=E1101
            except pywintypes.com_error:
                print(f"Etapa: {sap_etapa_material} - Asfalto frio já foi retirado.")

            if sap_material == '50000328' and '50000263' not in procura_lacre:
                # Remove o lacre e inclui o lacre novo, apenas se não tiver.
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
        sap_material = super().testa_material_sap(self.tb_materiais)
        hidro_instalado = self.hidro
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

                if self.hidro is not None:
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
                    "Não foi inserido hidro, "
                    + f"incluindo o informado: {cod_hidro_instalado}")
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", sap_etapa_material)
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                self.tb_materiais.modifyCell(ultima_linha_material, "QUANT", "1")
                self.tb_materiais.setCurrentCell(ultima_linha_material, "QUANT")
                ultima_linha_material = ultima_linha_material + 1
                hidro_adicionado = True  # Hidrômetro foi adicionado

class CorteRestabMaterial(Almoxarifado):
    '''Classe de materiais de religação e supressão.'''
    def __init__(self, int_num_lordem, hidro, operacao, identificador, tb_materiais):
        super().__init__(int_num_lordem, hidro, operacao, identificador)
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

class LigacaoAguaMaterial(Almoxarifado):
    '''Classe de materiais de Troca de Conexão de Ligação de Água.'''
    def __init__(self, int_num_lordem, hidro, operacao, identificador, tb_materiais):
        super().__init__(int_num_lordem, hidro, operacao, identificador)
        self.tb_materiais = tb_materiais
    def receita_troca_de_conexao_de_ligacao_de_agua(self):
        '''Padrão de materiais na classe Troca de Conexão de Ligação de Água.'''
        sap_material = super().testa_material_sap(self.tb_materiais)
        if sap_material is None:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona CONEXOES MET LIGACOES FEMEA DN 20.
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

            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material, "ETAPA", self.identificador[1]
            )
            # Adiciona REGISTRO METALICO RAMAL PREDIAL DN 20.
            self.tb_materiais.modifyCell(
                    ultima_linha_material, "MATERIAL", "30006747"
                )
            self.tb_materiais.modifyCell(
                ultima_linha_material, "QUANT", "1"
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material, "QUANT"
            )
            ultima_linha_material = ultima_linha_material + 1
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

            if '30002394' not in material_lista:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", self.identificador[1]
                )
                # Adiciona CONEXOES MET LIGACOES FEMEA DN 20.
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

            if '30006747' not in material_lista:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material, "ETAPA", self.identificador[1]
                )
                # Adiciona REGISTRO METALICO RAMAL PREDIAL DN 20.
                self.tb_materiais.modifyCell(
                        ultima_linha_material, "MATERIAL", "30006747"
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

class RedeAguaMaterial(Almoxarifado):
    '''Classe de materiais de CRA.'''
    def __init__(self, int_num_lordem, hidro, operacao, identificador, tb_materiais):
        super().__init__(int_num_lordem, hidro, operacao, identificador)
        self.tb_materiais = tb_materiais
    def receita_reparo_de_rede_de_agua(self):
        '''Padrão de materiais na classe CRA.'''
        sap_material = super().testa_material_sap(self.tb_materiais)
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

                if sap_material == '30005088':
                    self.tb_materiais.modifyCheckbox(
                        n_material, "ELIMINADO", True
                    )
                    self.tb_materiais.InsertRows(str(ultima_linha_material))
                    self.tb_materiais.modifyCell(
                        ultima_linha_material, "ETAPA", self.identificador[1]
                    )
                    # Adiciona o UNIAO P/TUBO PEAD DE 20 MM.
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

            # Materiais do Global.
            self.materiais_contratada(self.tb_materiais)

def materiais(int_num_lordem, hidro_instalado, operacao, identificador):
    '''Função dos materiais de acordo com a TSE pai.'''
    servico = Almoxarifado(int_num_lordem, hidro_instalado, operacao, identificador)
    tb_materiais = servico.aba_materiais()
    servico.inspecao(tb_materiais)
