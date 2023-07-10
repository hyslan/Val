# pagador.py
'''Módulo para precificar na aba itens de preço.'''
import sys
from servicos_executados import verifica_tse
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from unitarios import dicionario
from cesta import cesta_dicionario
session = connect_to_sap()
(
    lista,
    _,
    _,
    _,
    planilha,
    _,
    _,
    _,
    _,
    _,
    tb_contratada,
    tb_tse_un,
    tb_tse_rem_base,
    _,
    tb_tse_invest,
    *_,
) = load_worksheets()


def precificador(tse, corte, relig):
    '''Função para apontar os itens de preço e selecionar.'''
    tse.GetCellValue(0, "TSE")  # Saber qual TSE é
    (
        tse_temp,
        reposicao,
        num_tse_linhas,
        tse_proibida,
        identificador,
        etapa_reposicao,
        mae,
        list_chave_rb_despesa,
        chave_rb_investimento,
        chave_unitario,
        unitario_reposicao,
        rem_base_reposicao_union,
        reposicao_geral
    ) = verifica_tse(
        tse)
    etapa_rem_base = []
    etapa_unitario = []
    if tse_proibida is not None:
        print("TSE proibida de ser valorada.")
    else:
        print(f"TSE: {tse_temp}, Reposição inclusa ou não: {reposicao_geral}")
        print(f"Chave unitario: {chave_unitario}")
        print(f"Chave RB: {list_chave_rb_despesa}, {chave_rb_investimento}")
        if chave_unitario:  # Verifica se está no Conjunto Unitários
            # Aba Itens de preço
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI").select()
            print("****Processo de Precificação****")
            # pylint: disable=E1121
            dicionario.unitario(
                chave_unitario[0],
                corte,
                relig,
                chave_unitario[3],
                num_tse_linhas,
                chave_unitario[4],
                identificador
            )
            etapa_unitario.append(chave_unitario[0])

        if chave_rb_investimento:
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
            cesta_dicionario.cesta(
                chave_rb_investimento[3],
                chave_rb_investimento[4],
                chave_rb_investimento,
                mae
            )

        if list_chave_rb_despesa:
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
            for chave_rb_despesa in list_chave_rb_despesa:
                print(f"reosicao rb: {reposicao}")
                print(f"etapa rep rb: {etapa_reposicao}")
                cesta_dicionario.cesta(
                    chave_rb_despesa[3],
                    chave_rb_despesa[4],
                    chave_rb_despesa,
                    mae
                )

    return (
        tse_proibida,
        identificador,
        list_chave_rb_despesa,
        chave_rb_investimento,
        chave_unitario,
        etapa_rem_base,
        etapa_unitario
    )
