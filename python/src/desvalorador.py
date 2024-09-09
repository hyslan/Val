"""Módulo para desfazer valoração."""

from __future__ import annotations

import sys
import typing

import pywintypes
from tqdm import tqdm

from python.src.etl import pendentes_csv, pendentes_excel
from python.src.transact_zsbmm216 import Transacao

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch


def desvalorador(contrato: str, session: CDispatch) -> None:
    """Função desvalorador."""
    transacao = Transacao(contrato, "100", session)
    ask = input("É csv?")
    # * In case of CSV files
    if ask == "s":
        csv = pendentes_csv()
        limite_execucoes = len(csv)
        try:
            num_lordem = input("Insira o número da linha aqui: ")
            int_num_lordem = int(num_lordem)
            ordem = csv[int_num_lordem]
        except TypeError:
            sys.exit()

        # Loop para pagar as ordens da planilha do Excel
        for _ in tqdm(range(int_num_lordem, limite_execucoes), ncols=100):
            transacao.run_transacao(ordem)
            try:
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
                )
            # pylint: disable=E1101
            except pywintypes.com_error:
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
                continue

            try:
                session.findById("wnd[0]").SendVkey(2)
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
                ok = "Valoração desfeita com sucesso."
                rodape = session.findById("wnd[0]/sbar").Text  # Rodapé
                if ok == rodape:
                    pass
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
            # pylint: disable=E1101
            except pywintypes.com_error:
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
                continue
        # Fim
        return

    # * In case of Excel files
    pendentes_array = pendentes_excel()
    limite_execucoes = len(pendentes_array)
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = pendentes_array[int_num_lordem]
    except TypeError:
        return

    # Loop para pagar as ordens da planilha do Excel
    for ordem, cod_mun in tqdm(pendentes_array, ncols=100):
        # * Go to ZSBMM216 Transaction
        transacao.municipio = cod_mun
        transacao.run_transacao(ordem)
        try:
            session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell",
            )
        except pywintypes.com_error:
            # Incremento de Ordem.
            int_num_lordem += 1
            continue

        try:
            session.findById("wnd[0]").SendVkey(2)
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
            int_num_lordem += 1
        except pywintypes.com_error:
            # Incremento de Ordem.
            int_num_lordem += 1
            continue

    # Fim da desvaloração
