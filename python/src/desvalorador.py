"""Módulo para desfazer valoração."""
import sys
import pywintypes
from tqdm import tqdm
from python.src.transact_zsbmm216 import Transacao
from python.src.excel_tbs import load_worksheets
from python.src.etl import pendentes_csv

(
    lista,
    _,
    _,
    _,
    planilha,
    *_,
) = load_worksheets()


def desvalorador(contrato, session):
    """Função desvalorador"""
    transacao = Transacao(contrato, "100", session)
    ask = input("É csv?")
    if ask == "s":
        csv = pendentes_csv()
        limite_execucoes = len(csv)
        print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
        try:
            num_lordem = input("Insira o número da linha aqui: ")
            int_num_lordem = int(num_lordem)
            ordem = csv[int_num_lordem]
        except TypeError:
            print("Entrada inválida. Digite um número inteiro válido.")
            sys.exit()

        print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")

        # Loop para pagar as ordens da planilha do Excel
        for _ in tqdm(range(int_num_lordem, limite_execucoes), ncols=100):
            print(f"Linha atual: {int_num_lordem}.")
            print(f"Ordem atual: {ordem}")
            transacao.run_transacao(ordem)
            try:
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                )
            # pylint: disable=E1101
            except pywintypes.com_error:
                print(f"Ordem: {ordem} em medição definitiva.")
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
                continue

            try:
                session.findById("wnd[0]").SendVkey(2)
                print("Desvalorando!")
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
                ok = "Valoração desfeita com sucesso."
                rodape = session.findById("wnd[0]/sbar").Text  # Rodapé
                if ok == rodape:
                    print(f"{ordem} salva!")
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
            # pylint: disable=E1101
            except pywintypes.com_error:
                print(f"Ordem: {ordem} não valorada.")
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
                continue
        # Fim
        print("- Val: Desvaloração concluída.")
        return

    limite_execucoes = planilha.max_row
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = planilha.cell(row=int_num_lordem, column=1).value
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        return

    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")

    # Loop para pagar as ordens da planilha do Excel
    for _ in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
        print(f"Linha atual: {int_num_lordem}.")
        print(f"Ordem atual: {ordem}")
        transacao.run_transacao(ordem)
        try:
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
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
            lista.save('sheets/lista.xlsx')
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
            continue

    # Fim da desvaloração
    print("- Val: Desvalorador concluído.")
