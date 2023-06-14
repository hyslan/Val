# pai_dicionario.py
'''Módulo Dicionário Pai.'''
# Biblotecas
import sys
# Módulos Unitários
from tsepai.pai_unitario.pai_cavalete import m_cavalete
from tsepai.pai_unitario.pai_supressao import m_supressao
from tsepai.pai_unitario.pai_religacao import m_religacao
from tsepai.pai_unitario.pai_hidrometro import m_hidrometro
# Módulos Remuneração base
from tsepai.pai_cesta.pai_despesa.pai_cavalete import m_cavalete_rb


def pai_servico_unitario(servico_temp, n_etapa):
    '''Função condicional das chaves do dicionário unitário.'''

    dicionario_pai_unitario = {

        '142000': m_cavalete.Cavalete.RegularizarCv,
        '148000': m_cavalete.Cavalete.TrocaCvporUMA,
        '149000': m_cavalete.Cavalete.TrocaCvKit,
        '153000': m_cavalete.Cavalete.TrocaPeCvPrev,
        '201000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '203000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '203500': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '204000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '205000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '206000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '207000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        '215000': m_hidrometro.Hidrometro.UnitarioHidrometro,
        # '253000':
        # '254000':
        # '255000':
        # '262000':
        # '265000':
        # '266000':
        # '268000':
        # '269000':
        # '284500':
        # '286000':
        # '304000':
        '405000': m_supressao.Supressao.SuprimirLigacaodeAgua,
        '414000': m_supressao.Supressao.SuprimirLigacaodeAgua,
        '450500': m_religacao.Religacao.ReativadaLigacaodeAgua,
        '453000': m_religacao.Religacao.ReativadaLigacaodeAgua,
        '455500': m_religacao.Religacao.ReativadaLigacaodeAgua,
        '463000': m_religacao.Religacao.ReativadaLigacaodeAgua,
        '465000': m_religacao.Religacao.ReativadaLigacaodeAgua,
        '467500': m_religacao.Religacao.ReativadaLigacaodeAgua,
        '475500': m_religacao.Religacao.ReativadaLigacaodeAgua,
        # '502000':
        # '505000':
        # '506000':
        # '508000':
        # '537000':
        # '537100':
        # '538000':
        # '561000':
        # '565000':
        # '569000':
        # '581000':
        # '585000':
        # '713000':
        # '713500':
    }

    if servico_temp in dicionario_pai_unitario:
        print(
            f"TSE está inclusa no Dicionário de Pai Unitário: {servico_temp}")
        metodo = dicionario_pai_unitario[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao = metodo(n_etapa)
    else:
        print("TSE não Encontrada no Dicionário de Pai Unitário!")
        sys.exit()
    # Retorno
    return reposicao


def pai_servico_cesta(servico_temp):
    '''Função condicional das chaves do dicionário Remuneração Base.'''
    dicionario_pai_cesta = {

        '130000': m_cavalete_rb.Cavalete.reparo_cv,
        '140000': m_cavalete_rb.Cavalete.reparo_de_registro_de_cv,
        '140100': m_cavalete_rb.Cavalete.troca_de_registro_de_cv,
        # '283000':
        # '287000':
        # '321000':
        # '325000':
        # '332500':
        # '328000':
        # '330000':
        # '332000':
        # '416000':
        # '560000':
        # '567000':
        # '569000':
        # '580000':
        # '539000':
        # '540000':
        # '591000':
        # '284000':

    }

    if servico_temp in dicionario_pai_cesta:
        print(f"TSE está inclusa no Dicionário de Pai Cesta: {servico_temp}")
        metodo = dicionario_pai_cesta[servico_temp]
        metodo()  # Chama o método de uma classe dentro do Dicionário
    else:
        print("TSE não Encontrada no Dicionário de Pai Cesta!")
        sys.exit()
