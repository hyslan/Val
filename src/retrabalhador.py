"""Módulo dos retrabalho da valoração."""
import sys
import pywintypes
import rich
from tqdm import tqdm
from rich.console import Console
from src.transact_zsbmm216 import Transacao
from src.confere_os import consulta_os
from src.excel_tbs import load_worksheets
from src.sql_view import Tabela

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

console: rich.console.Console = Console()


def tag_n7(servico, session):
    num_tse_linhas = servico.RowCount
    for n_tse in range(0, num_tse_linhas):
        servico.modifyCell(n_tse, "PAGAR", "n")
        # Retrabalho
        servico.modifyCell(n_tse, "CODIGO", "7")

    servico.pressEnter()
    print("Salvando valoração!")
    session.findById("wnd[0]").sendVKey(11)
    session.findById("wnd[1]/usr/btnBUTTON_1").press()
    rodape = session.findById("wnd[0]/sbar").Text  # Rodapé

    return rodape

# consulta os só precisa de status_sistema, status_usuario


def retrabalho(contrato, session):
    """Função Retrabalhador: N7"""
    salvo = "Ajustes de valoração salvos com sucesso."
    transacao = Transacao(contrato, "100", session)
    revalorar = False
    sheet_csv = input("Deseja carregar a planilha xlsx ou CSV?\n")
    if sheet_csv:
        resposta = input("São Ordens desvaloradas?\n")
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
        for _ in tqdm(range(int_num_lordem, limite_execucoes + 1), ncols=100):
            print(f"Linha atual: {int_num_lordem}.")
            print(f"Ordem atual: {ordem}")
            transacao.run_transacao(ordem)
            try:
                servico = session.findById(
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

            if revalorar is False:
                try:
                    session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
                    grid_historico = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAO_NAPI:9040/cntlCC_AJUSTES/shellcont/shell")
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
                        + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                    )

            rodape = tag_n7(servico, session)
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

    # SQL values
    else:
        dt_inicio = input("Digite a data de início no formato Y-m-d:\n")
        dt_fim = input("Digite a data final no formato Y-m-d:\n")
        sql = Tabela("", "")
        pendentes_array = sql.retrabalho_search(dt_inicio, dt_fim)
        limite_execucoes = len(pendentes_array)
        if limite_execucoes == 0:
            return None, True
        print(
            f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
        if limite_execucoes == 0:
            return None, True

        with console.status("[bold blue]Trabalhando..."):
            for ordem, cod_mun, empresa in tqdm(pendentes_array, ncols=100):
                try:
                    # Setando valores para a transação
                    transacao.municipio = cod_mun
                    transacao.contrato = empresa
                    transacao.run_transacao(ordem)

                    console.print("Processo de Serviços Executados",
                                  style="bold red underline", justify="center")
                    try:
                        session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                            + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                        )
                    # pylint: disable=E1101
                    except pywintypes.com_error:
                        print(f"Ordem: {ordem} em medição definitiva.")
                        sql.ordem = ordem
                        sql.valorada(obs="Definitiva")
                        continue

                    try:
                        session.findById("wnd[0]").SendVkey(2)
                        console.print("Desvalorando!")
                        session.findById("wnd[1]/usr/btnBUTTON_1").press()
                        ok = "Valoração desfeita com sucesso."
                        rodape = session.findById("wnd[0]/sbar").Text  # Rodapé
                        if ok == rodape:
                            console.print(ok)
                            transacao.run_transacao(ordem)
                            servico = session.findById(
                                "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                                + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                            )

                            rodape = tag_n7(servico, session)
                            if rodape == salvo:
                                console.print("[italic green]Foi Salvo com sucesso! :rocket:")
                                sql.ordem = ordem
                                sql.valorada("SIM")
                            else:
                                console.print(f"Ordem: {ordem} não foi salva. :pouting_face:", style="italic yellow")
                                sql.ordem = ordem
                                sql.valorada(obs="Não foi salvo")

                    # pylint: disable=E1101
                    except pywintypes.com_error:
                        console.print(f"Ordem: {ordem} não valorada.")
                        servico = session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                            + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                        )
                        tag_n7(servico, session)

                    # Clean Duplicates from tb_Valoradas
                    sql.clean_duplicates()

                except Exception as e:
                    console.print(f"Erro dentro do For Retrabalhor: {e}")
                    console.print_exception()

    # Fim do retrabalhador
    print("- Val: Retrabalhador concluído.")
