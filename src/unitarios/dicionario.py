# dicionario.py
"""Módulo do Dicionário de Unitários"""
# Bibliotecas
import sys
from src.unitarios.hidrometro.m_hidrometro import Hidrometro
from src.unitarios.supressao.m_supressao import Corte
from src.unitarios.religacao.m_religacao import Religacao
from src.unitarios.cavalete.m_cavalete import Cavalete
from src.unitarios.poco.m_poco import Poco
from src.unitarios.ligacao_esgoto.m_ligacao_esgoto import LigacaoEsgoto
from src.unitarios.ligacao_agua.m_ligacao_agua import LigacaoAgua


def selecionar_tse(etapa, corte, relig, reposicao, num_tse_linhas,
                   etapa_reposicao, identificador, posicao_rede,
                   profundidade, session, preco):
    """Dicionário de chaves para etapas de unitário."""
    cavalete = Cavalete(etapa, corte, relig, reposicao, num_tse_linhas,
                        etapa_reposicao, identificador, posicao_rede,
                        profundidade, session, preco)
    hidrometro = Hidrometro(etapa, corte, relig, reposicao, num_tse_linhas,
                            etapa_reposicao, identificador, posicao_rede,
                            profundidade, session, preco)
    supressao = Corte(etapa, corte, relig, reposicao, num_tse_linhas,
                      etapa_reposicao, identificador, posicao_rede,
                      profundidade, session, preco)
    religacao = Religacao(etapa, corte, relig, reposicao, num_tse_linhas,
                          etapa_reposicao, identificador, posicao_rede,
                          profundidade, session, preco)
    poco = Poco(etapa, corte, relig, reposicao, num_tse_linhas,
                etapa_reposicao, identificador, posicao_rede,
                profundidade, session, preco)
    ligacao_agua = LigacaoAgua(etapa, corte, relig, reposicao, num_tse_linhas,
                               etapa_reposicao, identificador, posicao_rede,
                               profundidade, session, preco)
    ligacao_esgoto = LigacaoEsgoto(etapa, corte, relig, reposicao, num_tse_linhas,
                                   etapa_reposicao, identificador, posicao_rede,
                                   profundidade, session, preco)

    dicionario_un = {
        '134000': cavalete.instalado_lacre,
        '135000': cavalete.instalado_lacre,
        # '142000': m_cavalete.Cavalete.
        '148000': cavalete.troca_cv_por_uma,
        '149000': cavalete.troca_cv_kit,
        '153000': cavalete.troca_pe_cv_prev,
        '153500': cavalete.troca_pe_cv_prev,
        '201000': hidrometro.troca_de_hidro_corretivo,
        '202000': hidrometro.desinclinado_hidrometro,
        '203000': hidrometro.troca_de_hidro_corretivo,
        '203500': hidrometro.desinclinado_hidrometro,
        '204000': hidrometro.troca_de_hidro_corretivo,
        '205000': hidrometro.troca_de_hidro_corretivo,
        '206000': hidrometro.troca_de_hidro_corretivo,
        '207000': hidrometro.troca_de_hidro_corretivo,
        '215000': hidrometro.troca_de_hidro_preventiva_agendada,
        '254000': ligacao_agua.ligacao_agua,
        '262000': ligacao_agua.subst_agua,
        '263000': ligacao_agua.subst_agua,
        '265000': ligacao_agua.subst_agua,
        '280000': ligacao_agua.png,
        '284500': ligacao_agua.tra_nv,
        '286000': ligacao_agua.tra_prev,
        '312000': poco.niv_cx_parada,
        '322000': poco.troca_de_caixa_de_parada,
        '404200': supressao.supressao,
        '405000': supressao.supressao,
        '406000': supressao.supressao,
        '407000': supressao.supressao,
        '414000': supressao.supressao,
        '450500': religacao.restabelecida,
        '453000': religacao.restabelecida,
        '455500': religacao.restabelecida,
        '463000': religacao.restabelecida,
        '465000': religacao.restabelecida,
        '466500': religacao.restabelecida,
        '467500': religacao.restabelecida,
        '472000': religacao.restabelecida,
        '475500': religacao.restabelecida,
        '502000': ligacao_esgoto.ligacao_esgoto,
        '505000': ligacao_esgoto.ligacao_esgoto,
        '506000': ligacao_esgoto.ligacao_esgoto,
        '508000': ligacao_esgoto.ligacao_esgoto,
        '537000': poco.nivelamento,
        '537100': poco.nivelamento,
        '538000': poco.nivelamento,
        '565000': ligacao_esgoto.png,
        '569000': ligacao_esgoto.tre

    }

    if etapa in dicionario_un:
        print(
            f"Etapa está inclusa no Dicionário de Unitários: {etapa}")
        return dicionario_un[etapa]
        # Chama o método de uma classe dentro do Dicionário
    else:
        print(F"TSE {etapa} não Encontrada no Dicionário de Unitários!")
        sys.exit()
