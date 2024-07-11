"""Módulo de seletor de tse e seus ramos"""
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
PAGAR_SIM = "s"
CODIGO_DESPESA = "5"
CODIGO_INVESTIMENTO = "6"
PERTENCE_SERVICO_PRINCIPAL = "3"
SERVICO_N_EXISTE_CONTRATO = "10"


class Pai:
    """Classe Pai da árvore de TSEs"""

    def __init__(self, session) -> None:
        self.session = session

    def desobstrucao(self) -> tuple[list, None, str, list]:
        """Serviços de DD/DC, Lavagem, Televisionado"""
        etapa_reposicao = []
        tse_temp_reposicao = []
        print("Iniciando processo Pai de DD/DC, Lavagem e Televisionado")
        return tse_temp_reposicao, None, "desobstrucao", etapa_reposicao

    def get_shell(self) -> tuple[any, int]:
        """Adquire o controle do grid Serviço"""
        servico_temp = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell")
        num_tse_linhas = servico_temp.RowCount
        return servico_temp, num_tse_linhas

    def processar_servico_temp(self, servico_temp, num_tse_linhas: int, identificador: str):
        """Área do Loop"""
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
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_temp_unitario(self, servico_temp,
                                        num_tse_linhas: int, identificador: str):
        """Área do Loop"""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse, sap_tse in enumerate(range(num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_temp_cesta(self, servico_temp,
                                     num_tse_linhas: int,
                                     codigo_despesa: str,
                                     identificador: str) -> tuple[list, None, str, list]:
        """Área do Loop"""
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
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", codigo_despesa)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            # COMPACTAÇÃO E SELAGEM DA BASE
            if sap_tse in ('730600', '730700'):
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", codigo_despesa)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_reposicao_dependente(self, servico_temp,
                                               num_tse_linhas, identificador):
        """Processar serviços que acompanham juntos reposições no pagamento"""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse, sap_tse in enumerate(range(num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, None, identificador, etapa_reposicao


class Cesta(Pai):
    """Ramo de Rem Base"""

    def processar_servico_cesta(self, identificador, codigo_despesa):
        """Processar serviços da Cesta"""
        print(f"Iniciando processo Pai de {identificador}")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_cesta(servico_temp, num_tse_linhas,
                                                 codigo_despesa, identificador)

    def cavalete(self):
        """Serviços de cavalete"""
        return self.processar_servico_cesta("cavalete", PERTENCE_SERVICO_PRINCIPAL)

    def ligacao_agua(self):
        """Serviços de água"""
        return self.processar_servico_cesta("ligacao_agua", CODIGO_DESPESA)

    def suprimido_ramal_agua_abandonado(self):
        """Suprimido Ramal de água abandonado (RB) - TSE 416000,
        SUPRIMIDO RAMAL ANTERIOR - TSE 451000 e
        SUPRIMIDO RAMAL AGUA ABAND NÃO VISIVEL - TSE 416500"""
        return self.processar_servico_cesta("supressao", CODIGO_DESPESA)

    def reparo_ramal_agua(self):
        """Reparo de ramal de água (RB) - TSE 288000"""
        return self.processar_servico_cesta("reparo_ramal_agua", CODIGO_DESPESA)

    def ligacao_esgoto(self):
        """Reparo de ramal de esgoto (RB)"""
        return self.processar_servico_cesta("ligacao_esgoto", CODIGO_DESPESA)

    def poco(self):
        """Reconstruções/Reparos de poços (RB)"""
        return self.processar_servico_cesta("ligacao_esgoto", CODIGO_DESPESA)

    def rede_agua(self):
        """Reparo de Rede de Água (RB) - TSE 332000"""
        return self.processar_servico_cesta("rede_agua", CODIGO_DESPESA)

    def gaxeta(self):
        """Aperto de Gaxeta válvula de Rede de água (RB)"""
        return self.processar_servico_cesta("gaxeta", CODIGO_DESPESA)

    def chumbo_junta(self):
        """Rebatido Chumbo na Junta (RB) - TSE 330000"""
        return self.processar_servico_cesta("chumbo_junta", CODIGO_DESPESA)

    def valvula(self):
        """Troca de Válvula de Rede de água (RB) - TSE 325000"""
        return self.processar_servico_cesta("valvula", CODIGO_DESPESA)

    def rede_esgoto(self):
        """Reparo de Rede de Esgoto (RB) - TSE 580000"""
        return self.processar_servico_cesta("rede_esgoto", CODIGO_DESPESA)


class Investimento(Pai):
    """Ramo de Investimento (RB)"""

    def processar_servico_investimento(self, identificador, codigo_despesa):
        """Processar serviços de Investimento"""
        print(f"Iniciando processo Pai de {identificador}")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_cesta(servico_temp, num_tse_linhas,
                                                 codigo_despesa, identificador)

    def tra(self):
        """Troca de Ramal de Água (RB) - TSE 284000"""
        return self.processar_servico_investimento("tra", CODIGO_INVESTIMENTO)


class Sondagem(Pai):
    """Ramo de Sondagem Vala Seca"""

    def ligacao_agua(self):
        """Sondagem de Ramal de Água (RB) - TSE 283000"""
        print("Iniciando processo Pai de Sondagem de Ramal de Água")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "ligacao_agua")

    def rede_agua(self):
        """Sondagem de Ramal de Água (RB) - TSE 321000"""
        print("Iniciando processo Pai de Sondagem de Rede de Água")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "rede_agua")

    def ligacao_esgoto(self):
        """Sondagem de Ramal de Água (RB) - TSE 567000"""
        print("Iniciando processo Pai de Sondagem de Ramal de Esgoto")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "ligacao_esgoto")

    def rede_esgoto(self):
        """Sondagem de Ramal de Água (RB) - TSE 591000"""
        print("Iniciando processo Pai de Sondagem de Ramal de Esgoto")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "rede_esgoto")


class Unitario(Pai):
    """Ramo de serviços pagos unitariamente"""

    def processar_reposicao_sem_preco(self, servico_temp, num_tse_linhas, identificador):
        """Processar serviços que não tenham como pagar as reposições"""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse, sap_tse in enumerate(range(num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                if sap_tse == '732000':
                    servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                    servico_temp.modifyCell(
                        n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

            if sap_tse in ('170301', '749000', '758500', '740000'):
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", SERVICO_N_EXISTE_CONTRATO)

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_sem_bloquete(self, servico_temp,
                                       num_tse_linhas, identificador):
        """Processar serviços que não cotntemplam bloquete/muro"""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse, sap_tse in enumerate(range(num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                if sap_tse == '758000':
                    servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                    servico_temp.modifyCell(
                        n_tse, "CODIGO", SERVICO_N_EXISTE_CONTRATO)
                continue

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(
                    n_tse, "CODIGO", SERVICO_N_EXISTE_CONTRATO)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_unitario(self, identificador):
        """Processar serviços Unitários"""
        print(f"Iniciando processo Pai de {identificador}")
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_unitario(servico_temp, num_tse_linhas,
                                                    identificador)

    def cavalete(self):
        """Serviços de unitários de cavalete"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "cavalete")

    def lacre(self):
        """Instalados lacres avulsos."""
        return [], None, "cavalete", []

    def cavaletes_proibidos(self):
        """ Trocar Cavalete kit"""
        return [], "cavaletes proibidos", "cavalete", []

    def troca_cv_por_uma(self):
        """Troca de Cavalete por UMA - TSE 148000"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_reposicao_dependente(servico_temp, num_tse_linhas, "cavalete")

    def hidrometro(self):
        """Serviços unitários de hidrômetros"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "hidrometro")

    def hidrometro_alterar_capacidade(self):
        """Verificar esse serviço aí rapaz"""
        return [], "sei não ein", "desinclinado", []

    def desinclinado_hidrometro(self):
        """O hidrômetro mais barato"""
        return self.processar_servico_unitario("desinclinado")

    def ligacao_agua_avulsa(self):
        """Ligação de água simples avulsa"""
        return self.processar_servico_unitario("ligacao_agua_nova")

    def tra_nv_png_agua_subst_tra_prev(self):
        """TRA Não Visível, PNG Água,
        Substituição de Ligação água, TRA Preventivo"""
        return self.processar_servico_unitario("ligacao_agua")

    def ligacao_esgoto_avulsa(self):
        """Ligação de Esgoto"""
        return self.processar_servico_unitario("ligacao_esgoto")

    def png_esgoto(self):
        """PNG Obra Esgoto"""
        return self.processar_servico_unitario("png_esgoto")

    def tre(self):
        """Troca de Ramal de Esgoto"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_reposicao_sem_preco(
            servico_temp, num_tse_linhas, "ligacao_esgoto")

    def det_descoberto_nivelado_reg_cx_parada(self):
        """Descobrir, Trocar caixa de parada - TSE 322000"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_reposicao_sem_preco(
            servico_temp, num_tse_linhas, "cx_parada")

    def nivelamento_poco(self):
        """Nivelamento PI/PV/TL"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_reposicao_dependente(
            servico_temp, num_tse_linhas, "poço")

    def religacao(self):
        """Religação unitário
        TSEs: 405000, 414000, 450500, 453000,
        455500, 463000, 465000, 466500, 467500, 475500"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_sem_bloquete(
            servico_temp, num_tse_linhas, "religacao")

    def supressao(self):
        """Supressão unitário"""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_sem_bloquete(
            servico_temp, num_tse_linhas, "supressao")
