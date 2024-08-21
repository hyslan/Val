"""Módulo para desfazer valoração."""
import sys

import pywintypes
from tqdm import tqdm

from python.src.etl import pendentes_csv, pendentes_excel
from python.src.transact_zsbmm216 import Transacao


def desvalorador(contrato, session):
    """Função desvalorador"""
    transacao = Transacao(contrato, "100", session)
    ask = input("É csv?")
    # * In case of CSV files
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
                    + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
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

    # * In case of Excel files
    pendentes_array = pendentes_excel()
    limite_execucoes = len(pendentes_array)
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}\n")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = pendentes_array[int_num_lordem]
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        return

    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")

    # Loop para pagar as ordens da planilha do Excel
    for ordem, cod_mun in tqdm(pendentes_array, ncols=100):
        print(f"Linha atual: {int_num_lordem}.")
        print(f"Ordem atual: {ordem}")
        # * Go to ZSBMM216 Transaction
        transacao.municipio = cod_mun
        transacao.run_transacao(ordem)
        try:
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
            )
        except pywintypes.com_error:
            print(f"Ordem: {ordem} em medição definitiva.")
            # Incremento de Ordem.
            int_num_lordem += 1
            continue

        try:
            session.findById("wnd[0]").SendVkey(2)
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
            int_num_lordem += 1
            print(f"Ordem {ordem} desvalorada.")
        except pywintypes.com_error:
            print(f"Ordem: {ordem} não valorada.")
            # Incremento de Ordem.
            int_num_lordem += 1
            continue

    # Fim da desvaloração
    print("- Val: Desvalorador concluído.")
