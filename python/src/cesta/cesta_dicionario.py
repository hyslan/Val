# cesta_dicionario.py
"""Dicionário da Remuneração Base do Global."""

# Bibliotecas
import win32com.client as win32

from python.src.cesta.itens_naovinculados.m_itens_naovinculados import Modalidade


def cesta(
    reposicao: str,
    etapa_reposicao: str,
    identificador: tuple[str, str, str, str, str],
    mae: bool,
    session: win32.CDispatch,
) -> Modalidade:
    """Seletor de modalidade por Pai de Cesta."""
    rem_base = Modalidade(
        reposicao,
        etapa_reposicao,
        identificador,
        mae,
        session,
    )
    itens_nv: win32.CDispatch = rem_base.aba_nao_vinculados()
    rem_base.testa_modalidade_sap(itens_nv)
    rem_base.inspetor(itens_nv)
    return rem_base
