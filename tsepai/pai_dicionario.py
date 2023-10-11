# pai_dicionario.py
'''Módulo Dicionário Pai.'''
# Biblotecas
import sys
# Módulos Unitários
from tsepai.pai_unitario.pai_cavalete import m_cavalete
from tsepai.pai_unitario.pai_supressao import m_supressao
from tsepai.pai_unitario.pai_religacao import m_religacao
from tsepai.pai_unitario.pai_hidrometro import m_hidrometro
from tsepai.pai_unitario.pai_poco import m_poco
from tsepai.pai_unitario.pai_ligacaoesgoto import m_ligacao_esgoto_un
from tsepai.pai_unitario.pai_ligacaoagua import m_ligacao_agua_un
# Módulos Remuneração base
from tsepai.pai_cesta.pai_despesa.pai_cavalete import m_cavalete_rb
from tsepai.pai_cesta.pai_despesa.pai_ligacaoagua import m_ligacao_agua_rb
from tsepai.pai_cesta.pai_despesa.pai_redeagua import m_rede_agua_rb
from tsepai.pai_cesta.pai_despesa.pai_ligacaoesgoto import m_ligacao_esgoto_rb
from tsepai.pai_cesta.pai_despesa.pai_redeesgoto import m_rede_esgoto_rb
from tsepai.pai_cesta.pai_despesa.pai_poco import m_poco_rb
# Módulo Remuneração Base - Investimento
from tsepai.pai_cesta.pai_investimento import m_tra_rb

# Módulo Remuneração Base - Sondagem
from tsepai.pai_cesta.pai_sondagem import m_sondagem_rb


def preservacao_interferencia():
    '''Captador da tse preservação.'''
    tse_temp_reposicao = []
    tse_proibida = None
    identificador = "preservacao"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def transformacao_lig():
    '''Captador da tse Transformação'''
    tse_temp_reposicao = []
    tse_proibida = "Ramo Transformação"
    identificador = "transformacao"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def troca_de_ramal_agua_un():
    '''Captador da tse TRA.'''
    tse_temp_reposicao = []
    tse_proibida = "TRA"
    identificador = "tra"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def pai_servico_unitario(servico_temp):
    '''Função condicional das chaves do dicionário unitário.'''

    dicionario_pai_unitario = {
        # '134000': INSTALADO LACRE DIVERSOS
        # '135000': INSTALADO LACRE NUMERADO
        '138000': m_cavalete.Cavalete.readequado_cavalete,
        '142000': m_cavalete.Cavalete.regularizar_cv,
        '148000': m_cavalete.Cavalete.troca_cv_por_uma,
        '149000': m_cavalete.Cavalete.trocar_cv_kit,
        '153000': m_cavalete.Cavalete.troca_pe_cv_prev,
        '201000': m_hidrometro.Hidrometro.un_hidrometro,
        '202000': m_hidrometro.Hidrometro.desincl_hidrometro,
        '203000': m_hidrometro.Hidrometro.un_hidrometro,
        '203500': m_hidrometro.Hidrometro.desincl_hidrometro,
        '204000': m_hidrometro.Hidrometro.un_hidrometro,
        '205000': m_hidrometro.Hidrometro.un_hidrometro,
        '206000': m_hidrometro.Hidrometro.un_hidrometro,
        '207000': m_hidrometro.Hidrometro.un_hidrometro,
        '211000': m_hidrometro.Hidrometro.hidrometro_alterar_capacidade,
        '215000': m_hidrometro.Hidrometro.un_hidrometro,
        '253000': troca_de_ramal_agua_un,
        '254000': m_ligacao_agua_un.LigacaoAgua.ligacao_agua_avulsa,
        '255000': troca_de_ramal_agua_un,
        '262000': troca_de_ramal_agua_un,
        '265000': troca_de_ramal_agua_un,
        '266000': transformacao_lig,
        '267000': transformacao_lig,
        '268000': transformacao_lig,
        '269000': transformacao_lig,
        '280000': m_ligacao_agua_un.LigacaoAgua.png,
        '284500': m_ligacao_agua_un.LigacaoAgua.tra_nv,
        '286000': troca_de_ramal_agua_un,
        # '304000': DESCOBERTA VALVULA DE REDE DE AGUA
        '322000': m_poco.Poco.det_descoberto_nivelado_reg_cx_parada,
        '405000': m_supressao.Supressao.suprimir_ligacao_de_agua,
        '406000': m_supressao.Supressao.suprimir_ligacao_de_agua,
        '407000': m_supressao.Supressao.suprimir_ligacao_de_agua,
        '414000': m_supressao.Supressao.suprimir_ligacao_de_agua,
        '450500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '453000': m_religacao.Religacao.reativada_ligacao_de_agua,
        '455500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '463000': m_religacao.Religacao.reativada_ligacao_de_agua,
        '465000': m_religacao.Religacao.reativada_ligacao_de_agua,
        '467500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '475500': m_religacao.Religacao.reativada_ligacao_de_agua,
        '502000': m_ligacao_esgoto_un.LigacaoEsgoto.ligacao_esgoto_avulsa,
        '505000': m_ligacao_esgoto_un.LigacaoEsgoto.ligacao_esgoto_avulsa,
        '506000': m_ligacao_esgoto_un.LigacaoEsgoto.ligacao_esgoto_avulsa,
        '508000': m_ligacao_esgoto_un.LigacaoEsgoto.tre,
        '537000': m_poco.Poco.nivelamento,
        '537100': m_poco.Poco.nivelamento,
        '538000': m_poco.Poco.nivelamento,
        # '561000': DESOBSTRUIDO RAMAL DE ESGOTO
        '565000': m_ligacao_esgoto_un.LigacaoEsgoto.png,
        '569000': m_ligacao_esgoto_un.LigacaoEsgoto.tre,
        # '581000': DESOBSTRUIDA REDE DE ESGOTO
        # '585000': LAVAGEM DE REDE DE ESGOTO PREVENTIVA
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
        '283500': m_sondagem_rb.Sondagem.sondagem_de_ramal_de_agua,
        '284000': m_tra_rb.TrocaRamalAgua.troca_de_ramal_de_agua,
        '287000': m_ligacao_agua_rb.LigacaoAgua.troca_de_conexao_lig_agua,
        '288000': m_ligacao_agua_rb.LigacaoAgua.reparo_de_ramal_de_agua,
        '321000': m_sondagem_rb.Sondagem.sondagem_de_rede_de_agua,
        '321500': m_sondagem_rb.Sondagem.sondagem_de_rede_de_agua,
        # '325000':
        '328000': m_rede_agua_rb.RedeAgua.aperto_gaxeta_valvula,
        # '330000':
        '332000': m_rede_agua_rb.RedeAgua.reparo_de_rede_de_agua,
        '416000': m_ligacao_agua_rb.LigacaoAgua.suprimido_ramal_de_agua_abandonado,
        # '539000':
        # '540000':
        '560000': m_ligacao_esgoto_rb.LigacaoEsgoto.reparo_de_ramal_de_esgoto,
        '567000': m_sondagem_rb.Sondagem.sondagem_de_ramal_de_esgoto,
        # '569000':
        '580000': m_rede_esgoto_rb.RedeEsgoto.reparo_de_rede_de_esgoto,
        '539000': m_poco_rb.Poco.reconstruido_poco,
        '540000': m_poco_rb.Poco.reconstruido_poco,
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
