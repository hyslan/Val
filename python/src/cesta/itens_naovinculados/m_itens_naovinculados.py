# m_itens_naovinculados.py
"""Módulo para aba modalidade."""

import sys

import pywintypes
from rich.console import Console

console = Console()


class Modalidade:
    """Aba de Modalidade para Remuneração Base."""

    def __init__(self, reposicao, etapa_reposicao, identificador, mae, session) -> None:
        self.reposicao = reposicao
        self.etapa_reposicao = etapa_reposicao
        # O identificador é uma tupla com três variáveis:
        # 0: TSE ; 1: Etapa da TSE; 2: Família pro match case;
        self.identificador = identificador
        self.mae = mae
        self.session = session

    def aba_nao_vinculados(self):
        """Abrir aba de itens não vinculados."""
        console.print(
            "Processo de Itens não vinculados",
            style="bold red underline",
            justify="center",
        )
        self.session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
        itens_nv = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:"
            + "ZSBMM_VALORACAO_NAPI:9035/cntlCC_ITNS_NVINCRB/shellcont/shell",
        )
        return itens_nv

    def testa_modalidade_sap(self, itens_nv):
        """Tratamento de erro - Modalidade."""
        try:
            itens_nv.GetCellValue(0, "NUMERO_EXT")
            return
        # pylint: disable=E1101
        except pywintypes.com_error:
            sys.exit()

    def inspetor(self, itens_nv):
        """Selecionador de Pai Rem base."""
        match self.identificador[2]:
            case "supr_restab" | "supressao":
                self.testa_modalidade_sap(itens_nv)
                self.fech_e_reab_lig(itens_nv)
            case "ligacao_agua" | "cavalete" | "reparo_ramal_agua":
                self.testa_modalidade_sap(itens_nv)
                self.manut_lig_agua(itens_nv)
            case "rede_agua" | "gaxeta" | "chumbo_junta" | "valvula":
                self.testa_modalidade_sap(itens_nv)
                self.manut_rede_agua(itens_nv)
            case "ligacao_esgoto":
                self.testa_modalidade_sap(itens_nv)
                self.manut_lig_esg(itens_nv)
            case "rede_esgoto" | "poço":
                self.testa_modalidade_sap(itens_nv)
                self.manut_rede_esg(itens_nv)
            case "tra":
                self.testa_modalidade_sap(itens_nv)
                self.troca_de_ramal_de_agua(itens_nv)
            case "desobstrucao":
                self.testa_modalidade_sap(itens_nv)
                self.desobstrucao(itens_nv)
            case _:
                print("TSE não identificada.")
                print(f"TSE: {self.identificador[0]}")
                sys.exit()

    def fech_e_reab_lig(self, itens_nv):
        """Módulo Religação e Supressão - Modalidade.
        CÓDIGO: 327041 - PE, 327051 - SM
        327061 - IT, 327071 - AA
        """
        print(
            "Iniciando processo de Modalide - REM BASE - " + "MOD DESP FECH E REAB LIG ",
        )
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa == self.identificador[1]
                and sap_itens_nv
                in (
                    "327041",
                    "327051",  # MLG
                    "327061",
                    "327071",  # MLQ
                    "327081",
                    "327091",
                    "327101",
                    "327111",
                    "327121",
                    "327131",
                    "327141",
                    "327151",
                )
            ) or (
                sap_tse in self.reposicao
                and sap_etapa in self.etapa_reposicao
                and sap_itens_nv
                in (
                    "327041",
                    "327051",
                    "327061",
                    "327071",
                    "327081",
                    "327091",
                    "327101",
                    "327111",
                    "327121",
                    "327131",
                    "327141",
                    "327151",
                )
            ):  # MLN
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def manut_lig_esg(self, itens_nv):
        """Módulo Ramal de Esgoto - Rem Base.
        CÓDIGO: 327042 - PE ou 327052 - SM
        327062 - IT, 327072 - AA
        """
        print(
            "Iniciando processo de Modalidade - " + "REM BASE - MOD DESP MANUT LIG ESG",
        )
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa == self.identificador[1]
                and sap_itens_nv
                in (
                    "327042",
                    "327052",  # MLG
                    "327062",
                    "327072",  # MLQ
                    "327082",
                    "327092",
                    "327102",
                    "327112",
                    "327122",
                    "327132",
                    "327142",
                    "327152",
                )
            ) or (
                sap_tse in self.reposicao
                and sap_etapa in self.etapa_reposicao
                and sap_itens_nv
                in (
                    "327042",
                    "327052",
                    "327062",
                    "327072",
                    "327082",
                    "327092",
                    "327102",
                    "327112",
                    "327122",
                    "327132",
                    "327142",
                    "327152",
                )
            ):  # MLN
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def manut_rede_esg(self, itens_nv):
        """Módulo Despesa Rede de Esgoto - RB.
        CÓDIGO: 327043 - PE ou 327053 - SM
        327063 - IT, 327073 - AA
        """
        print("Iniciando processo de Modalidade - " + "REM BASE - DESP MANUT REDE ESG")
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa == self.identificador[1]
                and sap_itens_nv
                in (
                    "327043",
                    "327053",  # MLG
                    "327063",
                    "327073",  # MLQ
                    "327083",
                    "327093",
                    "327103",
                    "327113",
                    "327123",
                    "327133",
                    "327143",
                    "327153",
                )
            ):  # MLN
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
            elif (
                sap_tse in self.reposicao
                and sap_etapa in self.etapa_reposicao
                and sap_itens_nv
                in (
                    "327043",
                    "327053",
                    "327063",
                    "327073",
                    "327083",
                    "327093",
                    "327103",
                    "327113",
                    "327123",
                    "327133",
                    "327143",
                    "327153",
                )
            ):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def manut_rede_agua(self, itens_nv):
        """Módulo Despesa Rede de água - RB
        CÓDIGO: 327045 - PE ou 327055 - SM
        327065 - IT, 327075 - AA
        """
        print(
            "Iniciando processo de Modalidade - " + "REM BASE - MOD DESP RP MAN RD AGUA",
        )
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa == self.identificador[1]
                and sap_itens_nv
                in (
                    "327045",
                    "327055",  # MLG
                    "327065",
                    "327075",  # MLQ
                    "327085",
                    "327095",
                    "327105",
                    "327115",
                    "327125",
                    "327135",
                    "327145",
                    "327155",
                )
            ):  # MLN
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
            elif (
                sap_tse in self.reposicao
                and sap_etapa in self.etapa_reposicao
                and sap_itens_nv
                in (
                    "327045",
                    "327055",
                    "327065",
                    "327075",
                    "327085",
                    "327095",
                    "327105",
                    "327115",
                    "327125",
                    "327135",
                    "327145",
                    "327155",
                )
            ):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def manut_lig_agua(self, itens_nv):
        """Módulo de modalidade envolvendo ramal de água e cavalete.
        CÓDIGO: 327046 - PE OU 327056 - SM
        327066 - IT, 327076 - AA
        """
        print(
            "Iniciando processo de Modalidade - "
            + "REM BASE - MOD DESP RP MAN LIG AGUA",
        )
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa in self.identificador[1]
                and sap_itens_nv
                in (
                    "327046",
                    "327056",  # MLG
                    "327066",
                    "327076",  # MLQ
                    "327086",
                    "327096",
                    "327106",
                    "327116",
                    "327126",
                    "327136",
                    "327146",
                    "327156",
                )
            ):  # MLN
                itens_nv.ModifyCheckBox(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
            if (
                sap_tse in self.reposicao
                and sap_etapa in self.etapa_reposicao
                and sap_itens_nv
                in (
                    "327046",
                    "327056",
                    "327066",
                    "327076",
                    "327086",
                    "327096",
                    "327106",
                    "327116",
                    "327126",
                    "327136",
                    "327146",
                    "327156",
                )
                and self.mae is not True
            ):
                itens_nv.ModifyCheckBox(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def troca_de_ramal_de_agua(self, itens_nv):
        """Módulo Investimento TRA - RB.
        CÓDIGO: 327050 - PE ou 327060 - SM
        327070 - IT, 327080 - AA
        """
        print(
            "Iniciando processo de Modalidade - " + "REM BASE - MOD INVEST TR LIG AGUA",
        )
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa == self.identificador[1]
                and sap_itens_nv
                in (
                    "327050",
                    "327060",  # MLG
                    "327070",
                    "327080",  # MLQ
                    "327090",
                    "327110",
                    "327120",
                    "327130",
                    "327140",
                    "327150",
                    "327160",
                )
            ):  # MLN
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
            elif (
                sap_tse in self.reposicao
                and sap_etapa in self.etapa_reposicao
                and sap_itens_nv
                in (
                    "327050",
                    "327060",
                    "327070",
                    "327080",
                    "327090",
                    "327110",
                    "327120",
                    "327130",
                    "327140",
                    "327150",
                    "327160",
                )
            ):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def desobstrucao(self, itens_nv):
        """Módulo desobstrução para NORTESUL"""
        print("Iniciando processo de Modalidade - " + "REM. BASE DE DESOBSTRUÇÃO")
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(n_modalidade, "TSE")
            if (
                sap_tse == self.identificador[0]
                and sap_etapa == self.identificador[1]
                and sap_itens_nv
                in (
                    # MLG
                    "360247",
                    "360245",
                    # MLQ
                    "360256",  # ARTUR ALVIM
                    "360252",  # ITAQUERA
                    # MLN
                    "360274",  # SUZANO
                    "360276",  # POÁ
                    "360278",  # ITAQUA
                    "360280",  # FERRAZ
                    "360282",  # ARUJÁ
                )
            ):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
