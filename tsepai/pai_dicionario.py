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
from tsepai.pai_cesta.pai_despesa.pai_ligacaoagua import m_ligacao_agua_rb
# Módulo Remuneração Base - Investimento
from tsepai.pai_cesta.pai_investimento import m_tra_rb
from tsepai.pai_cesta.pai_despesa.pai_redeagua import m_rede_agua_rb
from tsepai.pai_cesta.pai_despesa.pai_ligacaoesgoto import m_ligacao_esgoto_rb
from tsepai.pai_cesta.pai_despesa.pai_redeesgoto import m_rede_esgoto_rb
# Módulo Remuneração Base - Sondagem
from tsepai.pai_cesta.pai_sondagem import m_sondagem_rb


def preservacao_interferencia():
    '''Captador da tse preservação.'''
    tse_temp_reposicao = []
    tse_proibida = None
    identificador = "preservacao"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def pai_servico_unitario(servico_temp):
    '''Função condicional das chaves do dicionário unitário.'''

    dicionario_pai_unitario = {

        '142000': m_cavalete.Cavalete.regularizar_cv,
        '148000': m_cavalete.Cavalete.troca_cv_por_uma,
        '149000': m_cavalete.Cavalete.trocar_cv_kit,
        '153000': m_cavalete.Cavalete.troca_pe_cv_prev,
        '201000': m_hidrometro.Hidrometro.un_hidrometro,
        '203000': m_hidrometro.Hidrometro.un_hidrometro,
        '203500': m_hidrometro.Hidrometro.un_hidrometro,
        '204000': m_hidrometro.Hidrometro.un_hidrometro,
        '205000': m_hidrometro.Hidrometro.un_hidrometro,
        '206000': m_hidrometro.Hidrometro.un_hidrometro,
        '207000': m_hidrometro.Hidrometro.un_hidrometro,
        '215000': m_hidrometro.Hidrometro.un_hidrometro,
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
        '405000': m_supressao.Supressao.suprimir_ligacao_de_agua,
        '414000': m_supressao.Supressao.suprimir_ligacao_de_agua,
        '450500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '453000': m_religacao.Religacao.reativada_ligacao_de_agua,
        '455500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '463000': m_religacao.Religacao.reativada_ligacao_de_agua,
        '465000': m_religacao.Religacao.reativada_ligacao_de_agua,
        '467500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '475500': m_religacao.Religacao.reativada_ligacao_de_agua,
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
        '713000': preservacao_interferencia,
        '713500': preservacao_interferencia,
    }

    if servico_temp in dicionario_pai_unitario:
        print(
            f"TSE está inclusa no Dicionário de Pai Unitário: {servico_temp}")
        metodo = dicionario_pai_unitario[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao, tse_proibida, identificador, etapa_reposicao = metodo()
    else:
        print("TSE não Encontrada no Dicionário de Pai Unitário!")
        sys.exit()
    # Retorno
    return reposicao, tse_proibida, identificador, etapa_reposicao


def pai_servico_cesta(servico_temp):
    '''Função condicional das chaves do dicionário Remuneração Base.'''
    dicionario_pai_cesta = {

        '130000': m_cavalete_rb.Cavalete.reparo_cv,
        '140000': m_cavalete_rb.Cavalete.reparo_de_registro_de_cv,
        '140100': m_cavalete_rb.Cavalete.troca_de_registro_de_cv,
        '283000': m_sondagem_rb.Sondagem.sondagem_de_ramal_de_agua,
        '284000': m_tra_rb.TrocaRamalAgua.troca_de_ramal_de_agua,
        '287000': m_ligacao_agua_rb.LigacaoAgua.troca_de_conexao_lig_agua,
        '321000': m_sondagem_rb.Sondagem.sondagem_de_rede_de_agua,
        # '325000':
        # '328000':
        # '330000':
        '332000': m_rede_agua_rb.RedeAgua.reparo_de_rede_de_agua,
        '416000': m_ligacao_agua_rb.LigacaoAgua.suprimido_ramal_de_agua_abandonado,
        # '539000':
        # '540000':
        '560000': m_ligacao_esgoto_rb.LigacaoEsgoto.reparo_de_ramal_de_esgoto,
        '567000': m_sondagem_rb.Sondagem.sondagem_de_ramal_de_esgoto,
        # '569000':
        '580000': m_rede_esgoto_rb.RedeEsgoto.reparo_de_rede_de_esgoto,
        # '539000':
        # '540000':
        '591000': m_sondagem_rb.Sondagem.sondagem_de_rede_de_esgoto,

    }

    if servico_temp in dicionario_pai_cesta:
        print(f"TSE está inclusa no Dicionário de Pai Cesta: {servico_temp}")
        metodo = dicionario_pai_cesta[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao, tse_proibida, identificador, etapa_reposicao = metodo()
    else:
        print("TSE não Encontrada no Dicionário de Pai Cesta!")
        sys.exit()

    return reposicao, tse_proibida, identificador, etapa_reposicao
