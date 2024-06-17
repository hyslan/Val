"""Módulo dos retrabalho da valoração."""
import sys
import pywintypes
from tqdm import tqdm
from src.transact_zsbmm216 import Transacao
from src.confere_os import consulta_os
from src.excel_tbs import load_worksheets


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


# consulta os só precisa de status_sistema, status_usuario


def retrabalho(contrato, session):
    """Função Retrabalhador"""
    transacao = Transacao(contrato, "100", session)
    revalorar = False
    resposta = input("São Ordens desvaloradas?")
    if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
        revalorar = True

    limite_execucoes = planilha.max_row
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = planilha.cell(row=int_num_lordem, column=1).value
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        sys.exit()

    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")

    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        ordem_obs = planilha.cell(row=int_num_lordem, column=4)
        print(f"Linha atual: {int_num_lordem}.")
        print(f"Ordem atual: {ordem}")
        transacao.run_transacao(ordem)
        try:
            servico = session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
            )
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(f"Ordem: {ordem} em medição definitiva.")
            ordem_obs = planilha.cell(row=int_num_lordem, column=4)
            ordem_obs.value = "MEDIÇÃO DEFINITIVA"
            lista.save('sheets/lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
            continue

        if revalorar is False:
            try:
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
                grid_historico = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9040/cntlCC_AJUSTES/shellcont/shell")
                data_valorado = grid_historico.GetCellValue(0, "DATA")
                if data_valorado is not None:
                    print(f"OS: {ordem} já valorada.")
                    print(f"Data: {data_valorado}")
                    ordem_obs = planilha.cell(row=int_num_lordem, column=4)
                    ordem_obs.value = "Já Salvo"
                    lista.save('sheets/lista.xlsx')
                    # Incremento de Ordem.
                    int_num_lordem += 1
                    ordem = planilha.cell(row=int_num_lordem, column=1).value
                    continue

            # pylint: disable=E1101
            except pywintypes.com_error:
                print("OS Livre para valorar.")
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS").select()
                servico = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
                )

        num_tse_linhas = servico.RowCount
        for n_tse in range(0, num_tse_linhas):
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Retrabalho
            servico.modifyCell(n_tse, "CODIGO", "7")

        servico.pressEnter()
        print("Salvando valoração!")
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[1]/usr/btnBUTTON_1").press()

        status_sistema, status_usuario, *_ = consulta_os(
            ordem, session, contrato)
        print("Verificando se Ordem foi valorada.")
        if status_usuario == "EXEC VALO":
            print(f"Status da Ordem: {status_sistema}, {status_usuario}")
            print("Foi Salvo com sucesso!")
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Salvo"
            lista.save('sheets/lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
        else:
            print(f"Ordem: {ordem} não foi salva.")
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Não"
            lista.save('sheets/lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value

    # Fim do retrabalhador
    print("- Val: Retrabalhador concluído.")
