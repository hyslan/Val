# cesta_dicionario.py
"""Dicionário da Remuneração Base do Global."""
# Bibliotecas
from python.src.cesta.itens_naovinculados.m_itens_naovinculados import Modalidade


def cesta(reposicao, etapa_reposicao, identificador, mae, session):
    """Seletor de modalidade por Pai de Cesta."""
    rem_base = Modalidade(
        reposicao,
        etapa_reposicao,
        identificador,
        mae,
        session
    )
    itens_nv = rem_base.aba_nao_vinculados()
    rem_base.testa_modalidade_sap(itens_nv)
    rem_base.inspetor(itens_nv)
    return rem_base
