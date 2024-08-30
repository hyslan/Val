# pai_dicionario.py
"""Módulo Dicionário Pai."""

# Biblotecas
from __future__ import annotations

import sys
import typing

from python.src.tsepai import pais

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


def oh_pai(session: CDispatch) -> pais.Pai:
    """Aleluia Irmãos."""
    return pais.Pai(session)


def preservacao_interferencia() -> tuple[list[str], None | str, str, list[str]]:
    """Captador da tse preservação."""
    tse_temp_reposicao = [""]
    tse_proibida = None
    identificador = "preservacao"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def transformacao_lig() -> tuple[list[str], None | str, str, list[str]]:
    """Captador da tse Transformação."""
    tse_temp_reposicao = []
    tse_proibida = "Ramo Transformação"
    identificador = "transformacao"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def troca_de_ramal_agua_un() -> tuple[list[str], None | str, str, list[str]]:
    """Captador da tse TRA."""
    tse_temp_reposicao = []
    tse_proibida = "TRA"
    identificador = "tra"
    etapa_reposicao = []
    return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao


def pai_servico_unitario(servico_temp: str, session: CDispatch) -> tuple[list[str], None | str, str, list[str]]:
    """Função condicional das chaves do dicionário unitário."""
    pai_unitario = pais.Unitario(session)

    dicionario_pai_unitario = {
        "134000": pai_unitario.lacre,
        "135000": pai_unitario.lacre,
        "138000": pai_unitario.cavaletes_proibidos,
        "142000": pai_unitario.cavaletes_proibidos,
        "148000": pai_unitario.troca_cv_por_uma,
        "149000": pai_unitario.cavaletes_proibidos,
        "153000": pai_unitario.cavalete,
        "153500": pai_unitario.cavalete,
        "201000": pai_unitario.hidrometro,
        "202000": pai_unitario.desinclinado_hidrometro,
        "203000": pai_unitario.hidrometro,
        "203500": pai_unitario.desinclinado_hidrometro,
        "204000": pai_unitario.hidrometro,
        "205000": pai_unitario.hidrometro,
        "206000": pai_unitario.hidrometro,
        "207000": pai_unitario.hidrometro,
        "211000": pai_unitario.hidrometro_alterar_capacidade,
        "215000": pai_unitario.hidrometro,
        "216000": pai_unitario.hidrometro,
        "253000": troca_de_ramal_agua_un,  # Inclusão
        "254000": pai_unitario.ligacao_agua_avulsa,
        "255000": troca_de_ramal_agua_un,  # Ligação Cv múltiplo
        "262000": pai_unitario.tra_nv_png_agua_subst_tra_prev,
        "263000": pai_unitario.tra_nv_png_agua_subst_tra_prev,
        "265000": pai_unitario.tra_nv_png_agua_subst_tra_prev,
        "266000": transformacao_lig,
        "267000": transformacao_lig,
        "268000": transformacao_lig,
        "269000": transformacao_lig,
        "280000": pai_unitario.tra_nv_png_agua_subst_tra_prev,
        "284500": pai_unitario.tra_nv_png_agua_subst_tra_prev,
        "286000": pai_unitario.tra_nv_png_agua_subst_tra_prev,
        "304000": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "312000": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "322000": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "404200": pai_unitario.supressao,
        "405000": pai_unitario.supressao,
        "406000": pai_unitario.supressao,
        "407000": pai_unitario.supressao,
        "414000": pai_unitario.supressao,
        "450500": pai_unitario.religacao,
        "453000": pai_unitario.religacao,
        "455500": pai_unitario.religacao,
        "463000": pai_unitario.religacao,
        "465000": pai_unitario.religacao,
        "466500": pai_unitario.religacao,
        "467500": pai_unitario.religacao,
        "472000": pai_unitario.religacao,
        "475500": pai_unitario.religacao,
        "502000": pai_unitario.ligacao_esgoto_avulsa,
        "505000": pai_unitario.ligacao_esgoto_avulsa,
        "506000": pai_unitario.ligacao_esgoto_avulsa,
        "507000": pai_unitario.ligacao_esgoto_avulsa,
        "508000": pai_unitario.tre,
        "534000": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "534100": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "534200": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "534300": pai_unitario.det_descoberto_nivelado_reg_cx_parada,
        "537000": pai_unitario.nivelamento_poco,
        "537100": pai_unitario.nivelamento_poco,
        "538000": pai_unitario.nivelamento_poco,
        "565000": pai_unitario.png_esgoto,
        "569000": pai_unitario.tre,
        "713000": preservacao_interferencia,
        "713500": preservacao_interferencia,
    }

    if servico_temp in dicionario_pai_unitario:
        metodo = dicionario_pai_unitario[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao, tse_proibida, identificador, etapa_reposicao = metodo()
    else:
        sys.exit()
    # Retorno
    return reposicao, tse_proibida, identificador, etapa_reposicao


def pai_servico_cesta(servico_temp: str, session: CDispatch) -> tuple[list[str], None | str, str, list[str]]:
    """Função condicional das chaves do dicionário Remuneração Base."""
    pai_cesta = pais.Cesta(session)
    pai_sondagem = pais.Sondagem(session)
    pai_invest = pais.Investimento(session)
    dicionario_pai_cesta = {
        "130000": pai_cesta.cavalete,
        "138000": pai_cesta.cavalete,
        "140000": pai_cesta.cavalete,
        "140100": pai_cesta.cavalete,
        "283000": pai_sondagem.ligacao_agua,
        "283500": pai_sondagem.ligacao_agua,
        "284000": pai_invest.tra,
        "287000": pai_cesta.ligacao_agua,
        "288000": pai_cesta.reparo_ramal_agua,
        "321000": pai_sondagem.rede_agua,
        "321500": pai_sondagem.rede_agua,
        "325000": pai_cesta.valvula,
        "328000": pai_cesta.gaxeta,
        "330000": pai_cesta.chumbo_junta,
        "332000": pai_cesta.rede_agua,
        "415000": pai_cesta.suprimido_ramal_agua_abandonado,
        "416000": pai_cesta.suprimido_ramal_agua_abandonado,
        "416500": pai_cesta.suprimido_ramal_agua_abandonado,
        "560000": pai_cesta.ligacao_esgoto,
        "567000": pai_sondagem.ligacao_esgoto,
        "569000": pai_cesta.rede_esgoto,
        "580000": pai_cesta.rede_esgoto,
        "531000": pai_cesta.poco,
        "531100": pai_cesta.poco,
        "539000": pai_cesta.poco,
        "540000": pai_cesta.poco,
        "591000": pai_sondagem.rede_esgoto,
    }

    if servico_temp in dicionario_pai_cesta:
        metodo = dicionario_pai_cesta[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao, tse_proibida, identificador, etapa_reposicao = metodo()
    else:
        sys.exit()

    return reposicao, tse_proibida, identificador, etapa_reposicao


def pai_servico_desobstrucao(servico_temp: str, session: CDispatch) -> tuple[list[str], None | str, str, list[str]]:
    """Agregador de TSE de contrato NORTE SUL.

    para serviços de DD e DC.
    """
    pai_desobstrucao = oh_pai(session)
    dicionario_pai_desobstrucao = {
        "561000": pai_desobstrucao.desobstrucao,
        "568000": pai_desobstrucao.desobstrucao,
        "581000": pai_desobstrucao.desobstrucao,
        "584000": pai_desobstrucao.desobstrucao,
        "585000": pai_desobstrucao.desobstrucao,
        "592000": pai_desobstrucao.desobstrucao,
        "717000": pai_desobstrucao.desobstrucao,
    }

    if servico_temp in dicionario_pai_desobstrucao:
        metodo = dicionario_pai_desobstrucao[servico_temp]
        # Chama o método de uma classe dentro do Dicionário
        reposicao, tse_proibida, identificador, etapa_reposicao = metodo()
    else:
        sys.exit()

    return reposicao, tse_proibida, identificador, etapa_reposicao
