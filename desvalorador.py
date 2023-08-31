'''Módulo para desfazer valoração.'''
import pywintypes
from sap_connection import connect_to_sap
from transact_zsbmm216 import novasp
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
    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        ordem_obs = planilha.cell(row=int_num_lordem, column=4)
        print(f"Linha atual: {int_num_lordem}.")
        print(f"Ordem atual: {ordem}")
        novasp(ordem)
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
        session.findById("wnd[0]").SendVkey(2)
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    # Fim da desvaloração
    resposta = input(
        "- Val: Deseja valorar agora? \n")
    if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
        val()
