'''Módulo dos retrabalho da valoração.'''
import pywintypes
from tqdm import tqdm
import salvacao
from sap_connection import connect_to_sap
from transact_zsbmm216 import novasp
from transact_zsbmm216 import recape
from excel_tbs import load_worksheets


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


def retrabalho():
    '''Função Retrabalhador'''
    session = connect_to_sap()
    limite_execucoes = planilha.max_row
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = planilha.cell(row=int_num_lordem, column=1).value
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        print("Reiniciando o programa...")
        retrabalho()

    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")
    qtd_ordem = 0  # Contador de ordens pagas.
    contrato_recape = "4600044782"
    contrato_novasp = "4600041302"
    contrato = input("- Val: Qual o contrato?\n")
    if contrato == contrato_recape or contrato in ("RECAPE", "recape"):
        transacao = recape
    elif contrato == contrato_novasp or contrato in ("NOVASP", "novasp"):
        transacao = novasp
    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        ordem_obs = planilha.cell(row=int_num_lordem, column=4)
        print(f"Linha atual: {int_num_lordem}.")
        print(f"Ordem atual: {ordem}")
        transacao(ordem)
        try:
            servico = session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
            )
            num_tse_linhas = servico.RowCount
            for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
                servico.modifyCell(n_tse, "PAGAR", "n")
                servico.modifyCell(n_tse, "CODIGO", "7")  # Retrabalho

            servico.pressEnter()
            qtd_ordem = salvacao.salvar(ordem, qtd_ordem)
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Salvo"
            lista.save('lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
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

    # Fim do retrabalhador
    print("- Val: Retrabalhador concluído.")