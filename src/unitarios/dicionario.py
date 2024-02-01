# dicionario.py
'''Módulo do Dicionário de Unitários'''
# Bibliotecas
import sys
from src.unitarios.hidrometro.m_hidrometro import Hidrometro
from src.unitarios.supressao import m_supressao
from src.unitarios.religacao import m_religacao
from src.unitarios.cavalete.m_cavalete import Cavalete
from src.unitarios.poco import m_poco
from src.unitarios.ligacao_esgoto import m_ligacao_esgoto
from src.unitarios.ligacao_agua import m_ligacao_agua


class Dicionario:
    '''Classe seletora para outras classes'''

    @staticmethod
    def selecionar_tse(etapa):
        '''Dicionário de chaves para etapas de unitário.'''
        dicionario_un = {
            '134000': Cavalete.instalado_lacre,
            '135000': Cavalete.instalado_lacre,
            # '142000': m_cavalete.Cavalete.
            '148000': Cavalete.troca_cv_por_uma,
            '149000': Cavalete.troca_cv_kit,
            '153000': Cavalete.troca_pe_cv_prev,
            '153500': Cavalete.troca_pe_cv_prev,
            '201000': Hidrometro.troca_de_hidro_corretivo,
            '202000': Hidrometro.desinclinado_hidrometro,
            '203000': Hidrometro.troca_de_hidro_corretivo,
            '203500': Hidrometro.desinclinado_hidrometro,
            '204000': Hidrometro.troca_de_hidro_corretivo,
            '205000': Hidrometro.troca_de_hidro_corretivo,
            '206000': Hidrometro.troca_de_hidro_corretivo,
            '207000': Hidrometro.troca_de_hidro_corretivo,
            '215000': Hidrometro.troca_de_hidro_preventiva_agendada,
            '254000': m_ligacao_agua.LigacaoAgua.ligacao_agua,
            '262000': m_ligacao_agua.LigacaoAgua.subst_agua,
            '263000': m_ligacao_agua.LigacaoAgua.subst_agua,
            '265000': m_ligacao_agua.LigacaoAgua.subst_agua,
            '280000': m_ligacao_agua.LigacaoAgua.png,
            '284500': m_ligacao_agua.LigacaoAgua.tra_nv,
            '286000': m_ligacao_agua.LigacaoAgua.tra_prev,
            '312000': m_poco.Poco.niv_cx_parada,
            '322000': m_poco.Poco.troca_de_caixa_de_parada,
            '404200': m_supressao.Corte.supressao,
            '405000': m_supressao.Corte.supressao,
            '406000': m_supressao.Corte.supressao,
            '407000': m_supressao.Corte.supressao,
            '414000': m_supressao.Corte.supressao,
            '450500': m_religacao.Religacao.restabelecida,
            '453000': m_religacao.Religacao.restabelecida,
            '455500': m_religacao.Religacao.restabelecida,
            '463000': m_religacao.Religacao.restabelecida,
            '465000': m_religacao.Religacao.restabelecida,
            '467500': m_religacao.Religacao.restabelecida,
            '475500': m_religacao.Religacao.restabelecida,
            '502000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
            '505000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
            '506000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
            '508000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
            '537000': m_poco.Poco.nivelamento,
            '537100': m_poco.Poco.nivelamento,
            '538000': m_poco.Poco.nivelamento,
            '565000': m_ligacao_esgoto.LigacaoEsgoto.png_esgoto,
            '569000': m_ligacao_esgoto.LigacaoEsgoto.tre

        }

        if etapa in dicionario_un:
            print(
                f"Etapa está inclusa no Dicionário de Unitários: {etapa}")
            return dicionario_un[etapa]
            # Chama o método de uma classe dentro do Dicionário
        else:
            print("TSE não Encontrada no Dicionário!")
            sys.exit()
