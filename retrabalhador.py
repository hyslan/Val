'''Módulo dos retrabalho da valoração.'''
import sys
import pywintypes
from tqdm import tqdm
from sap_connection import connect_to_sap
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


def consulta_os(n_os, contrato, unadm):
    '''ZSBPM020 para consulta'''
    session = connect_to_sap()
    session.StartTransaction("ZSBPM020")
    campo_os = session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW")
    campo_os.Text = n_os
    session.findById("wnd[0]/usr/txtS_CONTR-LOW").text = contrato
    session.findById("wnd[0]/usr/txtS_UN_ADM-LOW").text = unadm
    session.findById("wnd[0]").sendVKey(8)
    consulta = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    status_sistema = consulta.GetCellValue(0, "STTXT")
    status_usuario = consulta.GetCellValue(0, "USTXT")

    return status_sistema, status_usuario


def novasp(ordem):
    '''Contrato NOVASP'''
    session = connect_to_sap()
    print("Iniciando valoração.")
    session.StartTransaction("ZSBMM216")
    # Unidade Administrativa
    session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
    # Contrato NOVASP
    session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600041302"
    session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
    sap_ordem = session.findById(
        "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
    sap_ordem.Text = ordem
    session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8


def recape(ordem):
    '''Contrato NOVASP'''
    session = connect_to_sap()
    print("Iniciando valoração.")
    session.StartTransaction("ZSBMM216")
    # Unidade Administrativa
    session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
    # Contrato NOVASP
    session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600044782"
    session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
    sap_ordem = session.findById(
        "wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
    sap_ordem.Text = ordem
    session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8


def retrabalho():
    '''Função Retrabalhador'''
    revalorar = False
    resposta = input("São Ordens desvaloradas?")
    if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
        revalorar = True

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
    contrato_recape = "4600044782"
    contrato_novasp = "4600041302"
    contrato = input("- Val: Qual o contrato?\n")
    if contrato == contrato_recape or contrato in ("RECAPE", "recape"):
        transacao = recape
        contrato = contrato_recape
        unadm = "344"
    elif contrato == contrato_novasp or contrato in ("NOVASP", "novasp"):
        transacao = novasp
        contrato = contrato_novasp
        unadm = "344"
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
                    lista.save('lista.xlsx')
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
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Retrabalho
            servico.modifyCell(n_tse, "CODIGO", "7")

        servico.pressEnter()
        print("Salvando valoração!")
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[1]/usr/btnBUTTON_1").press()

        status_sistema, status_usuario = consulta_os(
            ordem, contrato, unadm)
        print("Verificando se Ordem foi valorada.")
        if status_usuario == "EXEC VALO":
            print(f"Status da Ordem: {status_sistema}, {status_usuario}")
            print("Foi Salvo com sucesso!")
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Salvo"
            lista.save('lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
        else:
            print(f"Ordem: {ordem} não foi salva.")
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Não"
            lista.save('lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value

    # Fim do retrabalhador
    print("- Val: Retrabalhador concluído.")
