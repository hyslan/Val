"""Módulo de seletor de tse e seus ramos."""

from __future__ import annotations

import typing

from python.src.excel_tbs import load_worksheets

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch

(
    *_,
    tb_tse_n3,
    tb_tse_reposicao,
    _,
) = load_worksheets()

PAGAR_NAO = "n"
PAGAR_SIM = "s"
CODIGO_DESPESA = "5"
CODIGO_INVESTIMENTO = "6"
PERTENCE_SERVICO_PRINCIPAL = "3"
SERVICO_N_EXISTE_CONTRATO = "10"


class Pai:
    """Classe Pai da árvore de TSEs."""

    def __init__(self, session: CDispatch) -> None:
        """Inicializador da classe Pai.

        Args:
        ----
            session (CDispatch): Sessão do SapGui.

        """
        self.session = session

    def desobstrucao(self) -> tuple[list, None, str, list]:
        """Serviços de DD/DC, Lavagem, Televisionado."""
        etapa_reposicao = []
        tse_temp_reposicao = []
        return tse_temp_reposicao, None, "desobstrucao", etapa_reposicao

    def get_shell(self) -> tuple[CDispatch, int]:
        """Adquire o controle do grid Serviço."""
        servico_temp = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:" + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
        )
        num_tse_linhas = servico_temp.RowCount
        return servico_temp, num_tse_linhas

    def processar_servico_temp(
        self,
        servico_temp: CDispatch,
        num_tse_linhas: int,
        identificador: str,
    ) -> tuple[list, None, str, list]:
        """Área do Loop."""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse in range(num_tse_linhas):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_n3:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_temp_unitario(
        self,
        servico_temp: CDispatch,
        num_tse_linhas: int,
        identificador: str,
    ) -> tuple[list, None, str, list]:
        """Área do Loop."""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse in range(num_tse_linhas):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_n3:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_temp_cesta(
        self,
        servico_temp: CDispatch,
        num_tse_linhas: int,
        codigo_despesa: str,
        identificador: str,
    ) -> tuple[list, None, str, list]:
        """Área do Loop."""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse in range(num_tse_linhas):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", codigo_despesa)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_n3:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            # COMPACTAÇÃO E SELAGEM DA BASE
            if sap_tse in ("730600", "730700"):
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", codigo_despesa)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_reposicao_dependente(
        self,
        servico_temp: CDispatch,
        num_tse_linhas: int,
        identificador: str,
    ) -> tuple[list, None, str, list]:
        """Processar serviços que acompanham juntos reposições no pagamento."""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse in range(num_tse_linhas):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            if sap_tse in tb_tse_n3:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

        return tse_temp_reposicao, None, identificador, etapa_reposicao


class Cesta(Pai):
    """Ramo de Rem Base."""

    def processar_servico_cesta(self, identificador: str, codigo_despesa: str) -> tuple[list, None, str, list]:
        """Processar serviços da Cesta."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_cesta(servico_temp, num_tse_linhas, codigo_despesa, identificador)

    def cavalete(self) -> tuple[list, None, str, list]:
        """Serviços de cavalete."""
        return self.processar_servico_cesta("cavalete", PERTENCE_SERVICO_PRINCIPAL)

    def ligacao_agua(self) -> tuple[list, None, str, list]:
        """Serviços de água."""
        return self.processar_servico_cesta("ligacao_agua", CODIGO_DESPESA)

    def suprimido_ramal_agua_abandonado(self) -> tuple[list, None, str, list]:
        """Suprimido Ramal de água abandonado (RB) - TSE 416000.

        SUPRIMIDO RAMAL ANTERIOR - TSE 451000 e
        SUPRIMIDO RAMAL AGUA ABAND NÃO VISIVEL - TSE 416500.
        """
        return self.processar_servico_cesta("supressao", CODIGO_DESPESA)

    def reparo_ramal_agua(self) -> tuple[list, None, str, list]:
        """Reparo de ramal de água (RB) - TSE 288000."""
        return self.processar_servico_cesta("reparo_ramal_agua", CODIGO_DESPESA)

    def ligacao_esgoto(self) -> tuple[list, None, str, list]:
        """Reparo de ramal de esgoto (RB)."""
        return self.processar_servico_cesta("ligacao_esgoto", CODIGO_DESPESA)

    def poco(self) -> tuple[list, None, str, list]:
        """Reconstruções/Reparos de poços (RB)."""
        return self.processar_servico_cesta("ligacao_esgoto", CODIGO_DESPESA)

    def rede_agua(self) -> tuple[list, None, str, list]:
        """Reparo de Rede de Água (RB) - TSE 332000."""
        return self.processar_servico_cesta("rede_agua", CODIGO_DESPESA)

    def gaxeta(self) -> tuple[list, None, str, list]:
        """Aperto de Gaxeta válvula de Rede de água (RB)."""
        return self.processar_servico_cesta("gaxeta", CODIGO_DESPESA)

    def chumbo_junta(self) -> tuple[list, None, str, list]:
        """Rebatido Chumbo na Junta (RB) - TSE 330000."""
        return self.processar_servico_cesta("chumbo_junta", CODIGO_DESPESA)

    def valvula(self) -> tuple[list, None, str, list]:
        """Troca de Válvula de Rede de água (RB) - TSE 325000."""
        return self.processar_servico_cesta("valvula", CODIGO_DESPESA)

    def rede_esgoto(self) -> tuple[list, None, str, list]:
        """Reparo de Rede de Esgoto (RB) - TSE 580000."""
        return self.processar_servico_cesta("rede_esgoto", CODIGO_DESPESA)


class Investimento(Pai):
    """Ramo de Investimento (RB)."""

    def processar_servico_investimento(self, identificador: str, codigo_despesa: str) -> tuple[list, None, str, list]:
        """Processar serviços de Investimento."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_cesta(servico_temp, num_tse_linhas, codigo_despesa, identificador)

    def tra(self) -> tuple[list, None, str, list]:
        """Troca de Ramal de Água (RB) - TSE 284000."""
        return self.processar_servico_investimento("tra", CODIGO_INVESTIMENTO)


class Sondagem(Pai):
    """Ramo de Sondagem Vala Seca."""

    def ligacao_agua(self) -> tuple[list, None, str, list]:
        """Sondagem de Ramal de Água (RB) - TSE 283000."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "ligacao_agua")

    def rede_agua(self) -> tuple[list, None, str, list]:
        """Sondagem de Ramal de Água (RB) - TSE 321000."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "rede_agua")

    def ligacao_esgoto(self) -> tuple[list, None, str, list]:
        """Sondagem de Ramal de Água (RB) - TSE 567000."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "ligacao_esgoto")

    def rede_esgoto(self) -> tuple[list, None, str, list]:
        """Sondagem de Ramal de Água (RB) - TSE 591000."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "rede_esgoto")


class Unitario(Pai):
    """Ramo de serviços pagos unitariamente."""

    def processar_reposicao_sem_preco(
        self,
        servico_temp: CDispatch,
        num_tse_linhas: int,
        identificador: str,
    ) -> tuple[list, None, str, list]:
        """Processar serviços que não tenham como pagar as reposições."""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse in range(num_tse_linhas):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                if sap_tse == "732000":
                    servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                    servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            if sap_tse in tb_tse_n3:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            # BLOQUETE
            if sap_tse in ("738000", "740000"):
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

            if sap_tse in ("170301", "749000", "758500", "740000"):
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", SERVICO_N_EXISTE_CONTRATO)

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_sem_bloquete(
        self,
        servico_temp: CDispatch,
        num_tse_linhas: int,
        identificador: str,
    ) -> tuple[list, None, str, list]:
        """Processar serviços que não cotntemplam bloquete/muro."""
        tse_temp_reposicao = []
        etapa_reposicao = []

        for n_tse in range(num_tse_linhas):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_SIM)
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                if sap_tse == "758000":
                    servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                    servico_temp.modifyCell(n_tse, "CODIGO", SERVICO_N_EXISTE_CONTRATO)
                continue

            if sap_tse in tb_tse_n3:
                servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
                servico_temp.modifyCell(n_tse, "CODIGO", PERTENCE_SERVICO_PRINCIPAL)
                continue

            """ -------- SERÁ PAGO LRP ESPECIAL
            # if sap_tse in ('738000', '740000'):
            #     servico_temp.modifyCell(n_tse, "PAGAR", PAGAR_NAO)
            #     servico_temp.modifyCell(
            #         n_tse, "CODIGO", SERVICO_N_EXISTE_CONTRATO)
            #     tse_temp_reposicao.append(sap_tse)
            #     etapa_reposicao.append(etapa)
            """

        return tse_temp_reposicao, None, identificador, etapa_reposicao

    def processar_servico_unitario(self, identificador: str) -> tuple[list, None, str, list]:
        """Processar serviços Unitários."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp_unitario(servico_temp, num_tse_linhas, identificador)

    def cavalete(self) -> tuple[list, None, str, list]:
        """Serviços de unitários de cavalete."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "cavalete")

    def lacre(self) -> tuple[list, None, str, list]:
        """Instalados lacres avulsos."""
        return [], None, "cavalete", []

    def cavaletes_proibidos(self) -> tuple[list, str, str, list]:
        """Trocar Cavalete kit."""
        return [], "cavaletes proibidos", "cavalete", []

    def troca_cv_por_uma(self) -> tuple[list, None, str, list]:
        """Troca de Cavalete por UMA - TSE 148000."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_reposicao_dependente(servico_temp, num_tse_linhas, "cavalete")

    def hidrometro(self) -> tuple[list, None, str, list]:
        """Serviços unitários de hidrômetros."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_temp(servico_temp, num_tse_linhas, "hidrometro")

    def hidrometro_alterar_capacidade(self) -> tuple[list, str, str, list]:
        """Verificar esse serviço aí rapaz."""
        return [], "sei não ein", "desinclinado", []

    def desinclinado_hidrometro(self) -> tuple[list, None, str, list]:
        """O hidrômetro mais barato."""
        return self.processar_servico_unitario("desinclinado")

    def ligacao_agua_avulsa(self) -> tuple[list, None, str, list]:
        """Ligação de água simples avulsa."""
        return self.processar_servico_unitario("ligacao_agua_nova")

    def tra_nv_png_agua_subst_tra_prev(self) -> tuple[list, None, str, list]:
        """TRA Não Visível, PNG Água.

        Substituição de Ligação água, TRA Preventivo.
        """
        return self.processar_servico_unitario("ligacao_agua")

    def ligacao_esgoto_avulsa(self) -> tuple[list, None, str, list]:
        """Ligação de Esgoto."""
        return self.processar_servico_unitario("ligacao_esgoto")

    def png_esgoto(self) -> tuple[list, None, str, list]:
        """PNG Obra Esgoto."""
        return self.processar_servico_unitario("png_esgoto")

    def tre(self) -> tuple[list, None, str, list]:
        """Troca de Ramal de Esgoto."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_reposicao_sem_preco(servico_temp, num_tse_linhas, "ligacao_esgoto")

    def det_descoberto_nivelado_reg_cx_parada(self) -> tuple[list, None, str, list]:
        """Descobrir, Trocar caixa de parada.

        DESCOBERTA VALVULA DE REDE DE AGUA
        TSE: 304000, 322000.
        """
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_reposicao_sem_preco(servico_temp, num_tse_linhas, "cx_parada")

    def nivelamento_poco(self) -> tuple[list, None, str, list]:
        """Nivelamento PI/PV/TL."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_reposicao_dependente(servico_temp, num_tse_linhas, "poço")

    def religacao(self) -> tuple[list, None, str, list]:
        """Religação unitário.

        TSEs: 405000, 414000, 450500, 453000,
        455500, 463000, 465000, 466500, 467500, 475500.
        """
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_sem_bloquete(servico_temp, num_tse_linhas, "religacao")

    def supressao(self) -> tuple[list, None, str, list]:
        """Supressão unitário."""
        servico_temp, num_tse_linhas = self.get_shell()
        return self.processar_servico_sem_bloquete(servico_temp, num_tse_linhas, "supressao")
