# dicionario.py
"""Módulo do Dicionário de Unitários."""

# Bibliotecas
from __future__ import annotations

import logging
import typing

from python.src.unitarios.cavalete.m_cavalete import Cavalete
from python.src.unitarios.hidrometro.m_hidrometro import Hidrometro
from python.src.unitarios.ligacao_agua.m_ligacao_agua import LigacaoAgua
from python.src.unitarios.ligacao_esgoto.m_ligacao_esgoto import LigacaoEsgoto
from python.src.unitarios.poco.m_poco import Poco
from python.src.unitarios.religacao.m_religacao import Religacao
from python.src.unitarios.supressao.m_supressao import Corte

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch

logger = logging.getLogger(__name__)


def selecionar_tse(
    etapa: str,
    corte: str,
    relig: str,
    reposicao: str,
    num_tse_linhas: int,
    etapa_reposicao: str,
    identificador: list[str],
    posicao_rede: str,
    profundidade: str,
    session: CDispatch,
    preco: CDispatch,
) -> typing.Callable | None:
    """Dicionário de chaves para etapas de unitário.

    Args:
    ----
        etapa (str): Etapa pai
        corte (str): Supressão
        relig (str): Religação
        reposicao (str): Serviço Complementar
        num_tse_linhas (int): Count de linhas
        etapa_reposicao (str): Etapa da reposição
        identificador (list[str]): TSE, Etapa, id match case do almoxarifado.py
        posicao_rede (str): Posição da Rede
        profundidade (str): Profundidade
        session (CDispatch): Sessão do SAP
        preco (CDispatch): GRID de preços do SAP

    Returns:
    -------
        Cavalete | Hidrometro | Corte | Religacao | Poco | LigacaoAgua | LigacaoEsgoto: Classe de Unitário Instanciada.

    """
    cavalete = Cavalete(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )
    hidrometro = Hidrometro(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )
    supressao = Corte(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )
    religacao = Religacao(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )
    poco = Poco(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )
    ligacao_agua = LigacaoAgua(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )
    ligacao_esgoto = LigacaoEsgoto(
        etapa,
        corte,
        relig,
        reposicao,
        num_tse_linhas,
        etapa_reposicao,
        identificador,
        posicao_rede,
        profundidade,
        session,
        preco,
    )

    dicionario_un = {
        "134000": cavalete.instalado_lacre,
        "135000": cavalete.instalado_lacre,
        # '142000': m_cavalete.Cavalete.
        "148000": cavalete.troca_cv_por_uma,
        "149000": cavalete.troca_cv_kit,
        "153000": cavalete.troca_pe_cv_prev,
        "153500": cavalete.troca_pe_cv_prev,
        "201000": hidrometro.troca_de_hidro_corretivo,
        "202000": hidrometro.desinclinado_hidrometro,
        "203000": hidrometro.troca_de_hidro_corretivo,
        "203500": hidrometro.desinclinado_hidrometro,
        "204000": hidrometro.troca_de_hidro_corretivo,
        "205000": hidrometro.troca_de_hidro_corretivo,
        "206000": hidrometro.troca_de_hidro_corretivo,
        "207000": hidrometro.troca_de_hidro_corretivo,
        "215000": hidrometro.troca_de_hidro_preventiva_agendada,
        "216000": hidrometro.troca_de_hidro_preventiva_agendada,
        "254000": ligacao_agua.ligacao_agua,
        "262000": ligacao_agua.subst_agua,
        "263000": ligacao_agua.subst_agua,
        "265000": ligacao_agua.subst_agua,
        "280000": ligacao_agua.png,
        "284500": ligacao_agua.tra_nv,
        "286000": ligacao_agua.tra_prev,
        "304000": poco.troca_de_caixa_de_parada,
        "312000": poco.niv_cx_parada,
        "322000": poco.troca_de_caixa_de_parada,
        "404200": supressao.supressao,
        "405000": supressao.supressao,
        "406000": supressao.supressao,
        "407000": supressao.supressao,
        "414000": supressao.supressao,
        "450500": religacao.restabelecida,
        "453000": religacao.restabelecida,
        "455500": religacao.restabelecida,
        "463000": religacao.restabelecida,
        "465000": religacao.restabelecida,
        "466500": religacao.restabelecida,
        "467500": religacao.restabelecida,
        "472000": religacao.restabelecida,
        "475500": religacao.restabelecida,
        "502000": ligacao_esgoto.ligacao_esgoto,
        "505000": ligacao_esgoto.ligacao_esgoto,
        "506000": ligacao_esgoto.ligacao_esgoto,
        "507000": ligacao_esgoto.ligacao_esgoto,
        "508000": ligacao_esgoto.ligacao_esgoto,
        "534000": poco.troca_de_caixa_de_parada,
        "534100": poco.troca_de_caixa_de_parada,
        # '534200': poco.troca_de_caixa_de_parada, DESCOBERTO E NIVELADO, It pays both?
        # '534300': poco.troca_de_caixa_de_parada, DESCOBERTO E NIVELADO, It pays both?
        "537000": poco.nivelamento,
        "537100": poco.nivelamento,
        "538000": poco.nivelamento,
        "565000": ligacao_esgoto.png,
        "569000": ligacao_esgoto.tre,
    }

    if etapa in dicionario_un:
        return dicionario_un[etapa]
        # Chama o método de uma classe dentro do Dicionário
    logger.error("Etapa %s não encontrada no dicionário.", etapa)
    return None
