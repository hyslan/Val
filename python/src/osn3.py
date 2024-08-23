"""Módulo dos N3 da valoração."""
import sys

import pywintypes
from tqdm import tqdm

from python.src.confere_os import consulta_os
from python.src.etl import pendentes_csv
from python.src.excel_tbs import load_worksheets
from python.src.transact_zsbmm216 import Transacao

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


def pertencedor(contrato, session) -> None:
    """Função N3."""
    transacao = Transacao(contrato, "100", session)
    revalorar = False
    resposta = input("São Ordens desvaloradas?")
    if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
        revalorar = True

    ask = input("É csv?")
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
        for num_lordem in tqdm(range(int_num_lordem, limite_execucoes), ncols=100):
            transacao.run_transacao(ordem)
            try:
                servico = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell",
                )
            # pylint: disable=E1101
            except pywintypes.com_error:
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
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
                        # Incremento de Ordem.
                        int_num_lordem += 1
                        ordem = csv[int_num_lordem]
                        continue

                # pylint: disable=E1101
                except pywintypes.com_error:
                    session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS").select()
                    servico = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell",
                    )

            num_tse_linhas = servico.RowCount
            for n_tse in range(num_tse_linhas):
                servico.modifyCell(n_tse, "PAGAR", "n")
                # Pertence ao serviço principal
                servico.modifyCell(n_tse, "CODIGO", "3")

            servico.pressEnter()
            session.findById("wnd[0]").sendVKey(11)
            session.findById("wnd[1]/usr/btnBUTTON_1").press()

            status_sistema, status_usuario, *_ = consulta_os(
                ordem, session, contrato)
            if status_usuario == "EXEC VALO":
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]
            else:
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = csv[int_num_lordem]

        # Fim dos N3
        return

    # Casos de planilha Excel, arrumar.
    limite_execucoes = planilha.max_row
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = planilha.cell(row=int_num_lordem, column=1).value
    except TypeError:
        sys.exit()


    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        ordem_obs = planilha.cell(row=int_num_lordem, column=4)
        transacao.run_transacao(ordem)
        try:
            servico = session.findById(
                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell",
            )
        # pylint: disable=E1101
        except pywintypes.com_error:
            ordem_obs = planilha.cell(row=int_num_lordem, column=4)
            ordem_obs.value = "MEDIÇÃO DEFINITIVA"
            lista.save("sheets/lista.xlsx")
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
                    ordem_obs = planilha.cell(row=int_num_lordem, column=4)
                    ordem_obs.value = "Já Salvo"
                    lista.save("sheets/lista.xlsx")
                    # Incremento de Ordem.
                    int_num_lordem += 1
                    ordem = planilha.cell(row=int_num_lordem, column=1).value
                    continue

            # pylint: disable=E1101
            except pywintypes.com_error:
                session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS").select()
                servico = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell",
                )

        num_tse_linhas = servico.RowCount
        for n_tse in range(num_tse_linhas):
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Pertence ao serviço principal
            servico.modifyCell(n_tse, "CODIGO", "3")

        servico.pressEnter()
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[1]/usr/btnBUTTON_1").press()

        status_sistema, status_usuario, *_ = consulta_os(
            ordem, session, contrato)
        if status_usuario == "EXEC VALO":
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Salvo"
            lista.save("sheets/lista.xlsx")
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value
        else:
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "Não"
            lista.save("sheets/lista.xlsx")
            # Incremento de Ordem.
            int_num_lordem += 1
            ordem = planilha.cell(row=int_num_lordem, column=1).value

    # Fim dos N3
