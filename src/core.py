# core.py
"""Coração da Val."""
# pylint: disable=W0611
import sys
import time
import pywintypes
from tqdm import tqdm
from rich.console import Console
from rich.panel import Panel
from src import sql_view
from src.sap import Sap
from src.transact_zsbmm216 import Transacao
from src.confere_os import consulta_os
from src.pagador import precificador
from src.almoxarifado import materiais
from src.salvacao import salvar
from src.temporizador import cronometro_val
from src.wms.consulta_estoque import estoque
from src.nazare_bugou import oxe


def rollback(sap) -> None:
    sap.encerrar_sap()
    # down_sap()
    # print("Reiniciando programa")


def val(pendentes_array, session, contrato, revalorar):
    """Sistema Val."""
    console = Console()
    sap = Sap()
    transacao = Transacao(contrato, "100", session)

    try:
        sessions = sap.listar_sessoes()
    # pylint: disable=E1101
    except pywintypes.com_error:
        return

    try:
        if not contrato == "4600043760":
            if not sessions.Count == 6:
                new_session = sap.criar_sessao(sessions)
                estoque_hj = estoque(new_session, sessions, contrato)
            else:
                estoque_hj = estoque(session, sessions, contrato)
    except:
        return

    limite_execucoes = len(pendentes_array)
    print(
        f"Quantidade de ordens incluídas na lista: {limite_execucoes}")

    with console.status("[bold blue]Trabalhando..."):
        # Variáveis de Status da Ordem
        valorada = "EXEC VALO" or "NEXE VALO"
        fechada = "LIB"
        qtd_ordem = 0  # Contador de ordens pagas.
        # Loop para pagar as ordens
        for ordem, cod_mun in tqdm(pendentes_array, ncols=100):
            try:
                start_time = time.time()  # Contador de tempo para valorar.
                print(f"Ordem atual: {ordem}")
                print("Verificando Status da Ordem.")
                # Função consulta de Ordem.
                print("Iniciando Consulta.")
                (status_sistema,
                 status_usuario,
                 corte,
                 relig,
                 posicao_rede,
                 profundidade,
                 hidro,
                 operacao,
                 diametro_ramal,
                 diametro_rede
                 ) = consulta_os(ordem, session, contrato)
                # Consulta Status da Ordem
                if not status_sistema == fechada:
                    print(f"OS: {ordem} aberta.")
                    continue

                if revalorar is False:
                    if status_usuario == valorada:
                        print(f"OS: {ordem} já valorada.")
                        ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
                        ja_valorado.valorada("SIM")
                        continue

                transacao.municipio = cod_mun
                transacao.run_transacao(ordem)

                console.print("Processo de Serviços Executados",
                              style="bold red underline", justify="center")
                try:
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                    )
                # pylint: disable=E1101
                except pywintypes.com_error:
                    print(f"Ordem: {ordem} em medição definitiva.")
                    ja_valorado = sql_view.Tabela(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(obs="Definitiva")
                    continue

                try:
                    if revalorar is False:
                        session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA").select()
                        grid_historico = session.findById(
                            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABA/ssubSUB_TAB:"
                            + "ZSBMM_VALORACAO_NAPI:9040/cntlCC_AJUSTES/shellcont/shell")
                        data_valorado = grid_historico.GetCellValue(
                            0, "DATA")
                        if data_valorado is not None:
                            print(f"OS: {ordem} já valorada.")
                            print(f"Data: {data_valorado}")
                            ja_valorado = sql_view.Tabela(
                                ordem=ordem, cod_tse="")
                            ja_valorado.valorada(obs="SIM")
                            continue

                # pylint: disable=E1101
                except pywintypes.com_error:
                    print("OS Livre para valorar.")
                    session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS").select()
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAO_NAPI:9010/cntlCC_SERVICO/shellcont/shell"
                    )

                # TSE e Aba Itens de preço
                (
                    tse_proibida,
                    list_chave_rb_despesa,
                    list_chave_unitario,
                    chave_rb_investimento,
                    chave_unitario,
                    ligacao_errada,
                    profundidade_errada
                ) = precificador(tse, corte, relig,
                                 posicao_rede, profundidade, contrato, session)
                # debug
                # sys.exit(0)
                if ligacao_errada is True:
                    ja_valorado = sql_view.Tabela(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(obs="Sem posição de rede.")
                    continue

                if profundidade_errada is True:
                    ja_valorado = sql_view.Tabela(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(
                        obs="Sem profundidade do ramal.")
                    continue

                # Se a TSE não estiver no escopo da Val, vai pular pra próxima OS.
                if tse_proibida is not None:
                    ja_valorado = sql_view.Tabela(
                        ordem=ordem, cod_tse="")
                    ja_valorado.valorada(obs="Num Pode")
                    continue

                # Aba Materiais

                # RB - Investimento
                if chave_rb_investimento:
                    materiais(
                        hidro,
                        chave_rb_investimento[1],  # Etapa atual da chave
                        chave_rb_investimento,
                        diametro_ramal,
                        diametro_rede,
                        contrato,
                        estoque_hj,
                        posicao_rede,
                        session)
                # RB - Despesa
                if list_chave_rb_despesa and not contrato == "4600043760":
                    for chave_rb_despesa in list_chave_rb_despesa:
                        materiais(
                            hidro,
                            chave_rb_despesa[1],  # Etapa atual da chave
                            chave_rb_despesa,
                            diametro_ramal,
                            diametro_rede,
                            contrato,
                            estoque_hj,
                            posicao_rede,
                            session)
                # Unitários
                if list_chave_unitario:
                    for chave_unitario in list_chave_unitario:
                        materiais(
                            hidro,
                            chave_unitario[1],  # Etapa atual da chave
                            chave_unitario,
                            diametro_ramal,
                            diametro_rede,
                            contrato,
                            estoque_hj,
                            posicao_rede,
                            session)

                # Fim dos materiais
                # sys.exit(0)
                # Salvar Ordem
                qtd_ordem, rodape = salvar(
                    ordem, qtd_ordem, contrato, session)
                salvo = "Ajustes de valoração salvos com sucesso."
                if not salvo == rodape:
                    console.print(f"Ordem: {ordem} não foi salva.", style="italic red")
                    console.print(f"[bold yellow]Motivo: {rodape}")
                    break
                # Fim do contador de valoração.
                cronometro_val(start_time, ordem)
                console.print(
                    Panel.fit(
                        f"Quantidade de ordens valoradas: {qtd_ordem}."),
                    style="italic yellow")

            # pylint: disable=E1101
            except Exception as errocritico:
                print(f"args do errocritico: {errocritico}")
                _, descricao, _, _ = errocritico.args
                match descricao:
                    case 'Falha catastrófica':
                        console.print("[bold red]SAPGUI has crashed. :fire:")
                        # rollback(sap)
                    case 'Falha na chamada de procedimento remoto.':
                        console.print("[bold red]SAPGUI has been finished strangely. :fire:")
                        # rollback(sap)
                    case 'O servidor RPC não está disponível.':
                        console.print("[bold red]SAPGUI was weirdly disconnected. :fire:")
                        # rollback(sap)
                    case _:
                        console.print(
                            "[bold red underline]Aconteceu um Erro com a Val!"
                            + f"\n Fatal Error: {errocritico}")
                        console.print_exception(show_locals=True)
                        oxe()
                        sys.exit()

        validador = True

    return ordem, validador
