# m_itens_naovinculados.py
'''Módulo para aba modalidade.'''
import sys
import pywintypes
from sap_connection import connect_to_sap
session = connect_to_sap()


class Modalidade:
    '''Aba de Modalidade para Remuneração Base.'''

    def __init__(self, tse_pai, etapa_pai, reposicao, etapa_reposicao, identificador) -> None:
        self.tse_pai = tse_pai
        self.etapa_pai = etapa_pai
        self.reposicao = reposicao
        self.etapa_reposicao = etapa_reposicao
        self.identificador = identificador

    def aba_nao_vinculados(self):
        '''Abrir aba de itens não vinculados.'''
        print("****Processo de Itens não vinculados****")
        session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
        itens_nv = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        return itens_nv

    def testa_modalidade_sap(self, itens_nv):
        '''Tratamento de erro - Modalidade.'''
        try:
            itens_nv.GetCellValue(0, "NUMERO_EXT")
            return
        # pylint: disable=E1101
        except pywintypes.com_error:
            sys.exit()

    def inspetor(self, itens_nv):
        '''Selecionador de Pai Rem base.'''
        match self.identificador:
            case "supr_restab":
                print(1)
                self.testa_modalidade_sap(itens_nv)
            case "ligacao_agua":
                print(2)
                self.testa_modalidade_sap(itens_nv)
            case "rede_agua":
                print(3)
                self.testa_modalidade_sap(itens_nv)
            case "ligacao_esgoto":
                print(4)
                self.testa_modalidade_sap(itens_nv)
            case "rede_esgoto":
                print(5)
                self.testa_modalidade_sap(itens_nv)
            case "tra":
                print("investimento")
                self.testa_modalidade_sap(itens_nv)
            case _:
                print("Pai não identificado.")
                sys.exit()

    def fech_e_reab_lig(self, itens_nv):
        '''Módulo Religação e Supressão - Modalidade.'''
        print("Iniciando processo de Modalide - REM BASE - "
              + "MOD DESP FECH E REAB LIG - CÓDIGO: 327041 - PE ou 327051 - SM")
        n_modalidade = 0
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(
                n_modalidade, "NUMERO_EXT")
            if sap_itens_nv == str(327041) or sap_itens_nv == str(327051):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
                break

    def manut_lig_esg(self, itens_nv):
        print("Iniciando processo de Modalidade - "
              + "REM BASE - MOD DESP MANUT LIG ESG - CÓDIGO: 7042")
        n_modalidade = 0

    def manut_rede_esg(self, itens_nv):
        print("Iniciando processo de Modalidade - "
              + "REM BASE - DESP MANUT REDE ESG - CÓDIGO: 7043")

    def manut_rede_agua(self, itens_nv):
        print("Iniciando processo de Modalidade - "
              + "REM BASE - MOD DESP RP MAN RD AGUA - CÓDIGO: 7045")

    def manut_lig_agua(self, itens_nv):
        '''Módulo de modalidade envolvendo ramal de água e cavalete.'''
        print("Iniciando processo de Modalidade - "
              + "REM BASE - MOD DESP RP MAN LIG AGUA - CÓDIGO: 327046 - PE OU 327056 - SM")
        n_modalidade = 0
        num_modalidade_linhas = itens_nv.RowCount
        for n_modalidade in range(num_modalidade_linhas):
            sap_itens_nv = itens_nv.GetCellValue(
                n_modalidade, "NUMERO_EXT")
            sap_etapa = itens_nv.GetCellValue(
                n_modalidade, "ETAPA")
            sap_tse = itens_nv.GetCellValue(
                n_modalidade, "TSE")
            if sap_tse == self.tse_pai and sap_etapa == self.etapa_pai \
                    and sap_itens_nv in ('327046', '327056'):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()
            elif sap_tse == self.reposicao and sap_etapa == self.etapa_reposicao \
                    and sap_itens_nv in ('327046', '327056'):
                itens_nv.modifyCell(n_modalidade, "MEDICAO", True)
                itens_nv.SetCurrentCell(n_modalidade, "MEDICAO")
                itens_nv.pressf4()

    def troca_de_ramal_de_agua(self, itens_nv, sap_modalidade):
        print("Iniciando processo de Modalidade - "
              + "REM BASE - MOD INVEST TR LIG AGUA - CÓDIGO: 7050")