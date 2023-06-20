# dicionario.py
'''Módulo do Dicionário de Unitários'''
# Bibliotecas
import sys
from unitarios.hidrometro import m_hidrometro
from unitarios.supressao import m_supressao
from unitarios.religacao import m_religacao
from unitarios.cavalete import m_cavalete


def unitario(etapa, corte, relig, reposicao, num_tse_linhas, etapa_reposicao):
    '''Dicionário de chaves para etapas de unitário.'''
    dicionario_un = {
        '201000': m_hidrometro.Hidrometro.THD_456901,
        '203000': m_hidrometro.Hidrometro.THD_456901,
        '203500': m_hidrometro.Hidrometro.HD_456022,
        '204000': m_hidrometro.Hidrometro.THD_456901,
        '205000': m_hidrometro.Hidrometro.THD_456901,
        '206000': m_hidrometro.Hidrometro.THD_456901,
        '207000': m_hidrometro.Hidrometro.THD_456901,
        '215000': m_hidrometro.Hidrometro.THDPrev_456902,
        '405000': m_supressao.Corte.supressao,
        '414000': m_supressao.Corte.supressao,
        '450500': m_religacao.Religacao.restabelecida,
        '453000': m_religacao.Religacao.restabelecida,
        '455500': m_religacao.Religacao.restabelecida,
        '463000': m_religacao.Religacao.restabelecida,
        '465000': m_religacao.Religacao.restabelecida,
        '467500': m_religacao.Religacao.restabelecida,
        '475500': m_religacao.Religacao.restabelecida,
        # '142000': m_cavalete.Cavalete.
        # '148000': m_cavalete.Cavalete.
        '149000': m_cavalete.Cavalete.TrocaCvKit,
        '153000': m_cavalete.Cavalete.TrocaPeCvPrev,


    }

    if etapa in dicionario_un:
        print(f"Etapa está inclusa no Dicionário de Unitários: {etapa}")
        metodo = dicionario_un[etapa]
        # Chama o método de uma classe dentro do Dicionário
        metodo(corte, relig, reposicao, num_tse_linhas, etapa_reposicao)
    else:
        print("TSE não Encontrada no Dicionário!")
        sys.exit()
