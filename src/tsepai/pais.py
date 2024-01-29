'''Módulo de seletor de tse e seus ramos'''
from src.excel_tbs import load_worksheets

(
    *_,
    tb_tse_PertenceAoServicoPrincipal,
    tb_tse_ServicoNaoExistenoContrato,
    tb_tse_reposicao,
    tb_tse_Retrabalho,
    tb_tse_Asfalto,
) = load_worksheets()

PAGAR_NAO = "n"
CODIGO_DESPESA = "5"
CODIGO_INVESTIMENTO = "6"
CODIGO_N3 = "3"


class Pai():
    '''Classe Pai da árvore de TSEs'''

    def __init__(self, session) -> None:
        self.session = session

    def desobstrucao(self):
        '''Serviços de DD/DC, Lavagem, Televisionado'''
        etapa_reposicao = []
        tse_temp_reposicao = []
        print("Iniciando processo Pai de DD/DC, Lavagem e Televisionado")
        return tse_temp_reposicao, None, "desobstrucao", etapa_reposicao

    def get_shell(self):
        '''Adquire o controle do grid Serviço'''
        servico_temp = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        num_tse_linhas = servico_temp.RowCount
        return servico_temp, num_tse_linhas

    def processar_servico_temp_cesta(self, servico_temp, num_tse_linhas,
                                     codigo_despesa, identificador):
        '''Área do Loop'''
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse, sap_tse in enumerate(range(num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", codigo_despesa)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", CODIGO_N3)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", codigo_despesa)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, None, identificador, etapa_reposicao


class Cesta(Pai):
    '''Ramo de Rem Base'''

    def processar_servico_cesta(self, identificador, codigo_despesa):
        '''Processar serviços da Cesta'''
        print(f"Iniciando processo Pai de {identificador}")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_cesta(servico_temp, num_tse_linhas,
                                                 codigo_despesa, identificador)

    def cavalete(self):
        '''Serviços de cavalete'''
        return self.processar_servico_cesta("cavalete", CODIGO_DESPESA)

    def ligacao_agua(self):
        '''Serviços de água'''
        return self.processar_servico_cesta("ligacao_agua", CODIGO_DESPESA)

    def suprimido_ramal_agua_abandonado(self):
        '''Suprimido Ramal de água abandonado (RB) - TSE 416000'''
        return self.processar_servico_cesta("supressao", CODIGO_DESPESA)

    def reparo_ramal_agua(self):
        '''Reparo de ramal de água (RB) - TSE 288000'''
        return self.processar_servico_cesta("reparo_ramal_agua", CODIGO_DESPESA)

    def ligacao_esgoto(self):
        '''Reparo de ramal de esgoto (RB)'''
        return self.processar_servico_cesta("ligacao_esgoto", CODIGO_DESPESA)

    def poco(self):
        '''Reconstruções de poços (RB)'''
        return self.processar_servico_cesta("ligacao_esgoto", CODIGO_DESPESA)

    def rede_agua(self):
        '''Reparo de Rede de Água (RB) - TSE 332000'''
        return self.processar_servico_cesta("rede_agua", CODIGO_DESPESA)

    def gaxeta(self):
        '''Aperto de Gaxeta válvula de Rede de água (RB)'''
        return self.processar_servico_cesta("gaxeta", CODIGO_DESPESA)

    def chumbo_junta(self):
        '''Rebatido Chumbo na Junta (RB) - TSE 330000'''
        return self.processar_servico_cesta("chumbo_junta", CODIGO_DESPESA)

    def valvula(self):
        '''Troca de Válvula de Rede de água (RB) - TSE 325000'''
        return self.processar_servico_cesta("valvula", CODIGO_DESPESA)

    def rede_esgoto(self):
        '''Reparo de Rede de Esgoto (RB) - TSE 580000'''
        return self.processar_servico_cesta("rede_esgoto", CODIGO_DESPESA)


class Investimento(Pai):
    '''Ramo de Investimento (RB)'''

    def processar_servico_investimento(self, identificador, codigo_despesa):
        '''Processar serviços de Investimento'''
        print(f"Iniciando processo Pai de {identificador}")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_cesta(servico_temp, num_tse_linhas,
                                                 codigo_despesa, identificador)

    def tra(self):
        '''Troca de Ramal de Água (RB) - TSE 284000'''
        return self.processar_servico_investimento("tra", CODIGO_INVESTIMENTO)


class Sondagem(Pai):
    '''Ramo de Sondagem Vala Seca'''

    def processar_servico_temp(self, servico_temp, num_tse_linhas, identificador):
        '''Área do Loop'''
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse, sap_tse in enumerate(range(num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", CODIGO_N3)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def ligacao_agua(self):
        '''Sondagem de Ramal de Água (RB) - TSE 283000'''
        print("Iniciando processo Pai de Sondagem de Ramal de Água")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "ligacao_agua")

    def rede_agua(self):
        '''Sondagem de Ramal de Água (RB) - TSE 321000'''
        print("Iniciando processo Pai de Sondagem de Rede de Água")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "rede_agua")

    def ligacao_esgoto(self):
        '''Sondagem de Ramal de Água (RB) - TSE 567000'''
        print("Iniciando processo Pai de Sondagem de Ramal de Esgoto")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "ligacao_esgoto")

    def rede_esgoto(self):
        '''Sondagem de Ramal de Água (RB) - TSE 591000'''
        print("Iniciando processo Pai de Sondagem de Ramal de Esgoto")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "rede_esgoto")
