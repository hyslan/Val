# pagador.py
'''Módulo para precificar na aba itens de preço.'''
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
        chave_rb_despesa,
        chave_rb_investimento,
        chave_unitario
    ) = verifica_tse(
        tse)
    etapa_rem_base = []
    etapa_unitario = []
    if tse_proibida is not None:
        print("TSE proibida de ser valorada.")
    else:
        print(f"TSE: {tse_temp}, Reposição inclusa ou não: {reposicao}")
        for etapa in tse_temp:
            print(f"TSE selecionada para pagar: {etapa}")
            if etapa in tb_tse_un:  # Verifica se está no Conjunto Unitários
                print(f"{etapa} é unitário!")
                # Aba Itens de preço
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI").select()
                print("****Processo de Precificação****")
                # pylint: disable=E1121
                dicionario.unitario(
                    etapa,
                    corte,
                    relig,
                    reposicao,
                    num_tse_linhas,
                    etapa_reposicao,
                    identificador
                )
                etapa_unitario.append(etapa)

            elif etapa in tb_tse_rem_base or etapa in tb_tse_invest:
                print(f"{etapa} é Remuneração Base!")
                etapa_rem_base.append(etapa)

        if etapa_rem_base:
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
            for etapa_rb in etapa_rem_base:
                if etapa_rb in tb_tse_rem_base:
                    cesta_dicionario.cesta(
                        reposicao,
                        etapa_reposicao,
                        chave_rb_despesa,
                        mae
                    )
                else:
                    cesta_dicionario.cesta(
                        reposicao,
                        etapa_reposicao,
                        chave_rb_investimento,
                        mae
                    )

    return (
        tse_proibida,
        identificador,
        chave_rb_despesa,
        chave_rb_investimento,
        chave_unitario,
        etapa_rem_base,
        etapa_unitario
    )
