'''Módulo para desfazer valoração.'''
import pywintypes
from sap_connection import connect_to_sap
from transact_zsbmm216 import novasp
from transact_zsbmm216 import recape
from excel_tbs import load_worksheets
from tqdm import tqdm
from core import val

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
    _,
    tb_tse_un,
    tb_tse_rem_base,
    _,
    tb_tse_invest,
    *_,
) = load_worksheets()


def desvalorador():
    '''Função desvalorador'''
    limite_execucoes = planilha.max_row
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = planilha.cell(row=int_num_lordem, column=1).value
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        print("Reiniciando o programa...")
        desvalorador()
    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")
    contrato_recape = "4600044782"
    contrato_novasp = "4600041302"
    contrato = input("- Val: Qual o contrato?\n")
    if contrato == contrato_recape:
        transacao = recape
    elif contrato == contrato_novasp:
        transacao = novasp
    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        ordem_obs = planilha.cell(row=int_num_lordem, column=4)
        print(f"Linha atual: {int_num_lordem}.")
        print(f"Ordem atual: {ordem}")
        transacao(ordem)
        try:
            tse = session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
            )
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(f"Ordem: {ordem} em medição definitiva.")
            ordem_obs = planilha.cell(row=int_num_lordem, column=4)
            ordem_obs.value = "MEDIÇÃO DEFINITIVA"
            lista.save('lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
            continue

        try:
            session.findById("wnd[0]").SendVkey(2)
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(f"Ordem: {ordem} não valorada.")
            ordem_obs = planilha.cell(row=int_num_lordem, column=4)
            ordem_obs.value = "Não Valorada."
            lista.save('lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
            continue
    # Fim da desvaloração
    print("- Val: Desvalorador concluído.")
