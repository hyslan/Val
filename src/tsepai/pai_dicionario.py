# pai_dicionario.py
'''Módulo Dicionário Pai.'''
# Biblotecas
import sys
# Módulos Unitários
from src.tsepai.pai_unitario.pai_cavalete import m_cavalete
from src.tsepai.pai_unitario.pai_supressao import m_supressao
from src.tsepai.pai_unitario.pai_religacao import m_religacao
from src.tsepai.pai_unitario.pai_hidrometro import m_hidrometro
from src.tsepai.pai_unitario.pai_poco import m_poco
from src.tsepai.pai_unitario.pai_ligacaoesgoto import m_ligacao_esgoto_un
from src.tsepai.pai_unitario.pai_ligacaoagua import m_ligacao_agua_un
from src.tsepai import pais


def oh_pai(session):
    '''Aleluia Irmãos'''
    pai = pais.Pai(session)
    return pai


def preservacao_interferencia():
    '''Captador da tse preservação.'''
    tse_temp_reposicao = ['']
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
        '134000': m_cavalete.Cavalete.instalado_lacre,
        '135000': m_cavalete.Cavalete.instalado_lacre,
        '138000': m_cavalete.Cavalete.readequado_cavalete,
        '142000': m_cavalete.Cavalete.regularizar_cv,
        '148000': m_cavalete.Cavalete.troca_cv_por_uma,
        '149000': m_cavalete.Cavalete.trocar_cv_kit,
        '153000': m_cavalete.Cavalete.troca_pe_cv_prev,
        '153500': m_cavalete.Cavalete.troca_pe_cv_prev,
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
        '253000': troca_de_ramal_agua_un,  # Inclusão
        '254000': m_ligacao_agua_un.LigacaoAgua.ligacao_agua_avulsa,
        '255000': troca_de_ramal_agua_un,  # Ligação Cv múltiplo
        '262000': m_ligacao_agua_un.LigacaoAgua.tra_nv_png_subst_tra_prev,
        '263000': m_ligacao_agua_un.LigacaoAgua.tra_nv_png_subst_tra_prev,
        '265000': m_ligacao_agua_un.LigacaoAgua.tra_nv_png_subst_tra_prev,
        '266000': transformacao_lig,
        '267000': transformacao_lig,
        '268000': transformacao_lig,
        '269000': transformacao_lig,
        '280000': m_ligacao_agua_un.LigacaoAgua.tra_nv_png_subst_tra_prev,
        '284500': m_ligacao_agua_un.LigacaoAgua.tra_nv_png_subst_tra_prev,
        '286000': m_ligacao_agua_un.LigacaoAgua.tra_nv_png_subst_tra_prev,
        # '304000': DESCOBERTA VALVULA DE REDE DE AGUA
        '312000': m_poco.Poco.det_descoberto_nivelado_reg_cx_parada,
        '322000': m_poco.Poco.det_descoberto_nivelado_reg_cx_parada,
        '404200': m_supressao.Supressao.suprimir_ligacao_de_agua,
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


def pai_servico_cesta(servico_temp, session):
    '''Função condicional das chaves do dicionário Remuneração Base.'''
    pai_cesta = pais.Cesta(oh_pai(session))
    pai_sondagem = pais.Sondagem(oh_pai(session))
    pai_invest = pais.Investimento(oh_pai(session))
    dicionario_pai_cesta = {

        '130000': pai_cesta.cavalete,
        '138000': pai_cesta.cavalete,
        '140000': pai_cesta.cavalete,
        '140100': pai_cesta.cavalete,
        '283000': pai_sondagem.ligacao_agua,
        '283500': pai_sondagem.ligacao_agua,
        '284000': pai_invest.tra,
        '287000': pai_cesta.ligacao_agua,
        '288000': pai_cesta.reparo_ramal_agua,
        '321000': pai_sondagem.rede_agua,
        '321500': pai_sondagem.rede_agua,
        '325000': pai_cesta.valvula,
        '328000': pai_cesta.gaxeta,
        '330000': pai_cesta.chumbo_junta,
        '332000': pai_cesta.rede_agua,
        '416000': pai_cesta.suprimido_ramal_agua_abandonado,
        '560000': pai_cesta.ligacao_esgoto,
        '567000': pai_sondagem.ligacao_esgoto,
        '569000': pai_cesta.rede_esgoto,
        '580000': pai_cesta.rede_esgoto,
        '539000': pai_cesta.poco,
        '540000': pai_cesta.poco,
        '591000': pai_sondagem.rede_esgoto,

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


def pai_servico_desobstrucao(servico_temp, session):
    '''Agregador de TSE de contratos(ex: NORTE SUL)
    para serviços de DD e DC'''
    pai_desobstrucao = oh_pai(session)
    dicionario_pai_desobstrucao = {
        '561000': pai_desobstrucao.desobstrucao,
        '568000': pai_desobstrucao.desobstrucao,
        '581000': pai_desobstrucao.desobstrucao,
        '584000': pai_desobstrucao.desobstrucao,
        '585000': pai_desobstrucao.desobstrucao,
        '592000': pai_desobstrucao.desobstrucao,
        '717000': pai_desobstrucao.desobstrucao,

    }

    if servico_temp in dicionario_pai_desobstrucao:
        print(
            f"TSE está inclusa no Dicionário de Pai Desobstrução: {servico_temp}")
        metodo = dicionario_pai_desobstrucao[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao, tse_proibida, identificador, etapa_reposicao = metodo()
    else:
        print("TSE não Encontrada no Dicionário de Pai Desobstrução!")
        sys.exit()

    return reposicao, tse_proibida, identificador, etapa_reposicao
