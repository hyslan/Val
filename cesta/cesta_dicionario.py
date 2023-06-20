# cesta_dicionario.py
'''Dicionário da Remuneração Base do Global.'''
# Bibliotecas
import sys
from cesta.itens_naovinculados.m_itens_naovinculados import Modalidade


def cesta(etapa, reposicao, etapa_reposicao, identificador, etapa_pai):
    '''Seletor de modalidade por Pai de Cesta.'''
    rem_base = Modalidade(etapa,
                          etapa_pai,
                          reposicao,
                          etapa_reposicao,
                          identificador
                          )
    itens_nv = rem_base.aba_nao_vinculados()
    rem_base.testa_modalidade_sap(itens_nv)
    rem_base.inspetor(itens_nv)
