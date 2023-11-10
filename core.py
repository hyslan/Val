# core.py
'''Coração da Val.'''
# pylint: disable=W0611
import sys
import time
import pywintypes
from tqdm import tqdm
import sql_view
from sap_connection import connect_to_sap
from confere_os import consulta_os
from transact_zsbmm216 import novasp
from transact_zsbmm216 import gbitaquera
from pagador import precificador
from almoxarifado import materiais
from salvacao import salvar
from temporizador import cronometro_val
from sapador import down_sap


def val(pendentes_list, contrato, unadm):
    '''Sistema Val.'''
    validador = False
    input("- Val: Pressione Enter para iniciar...")
    limite_execucoes = len(pendentes_list)
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = pendentes_list[int_num_lordem]
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = pendentes_list[int_num_lordem]

    # Variáveis de Status da Ordem
    valorada = "EXEC VALO" or "NEXE VALO"
    fechada = "LIB"
    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")
    qtd_ordem = 0  # Contador de ordens pagas.
    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes-1), ncols=100):
        try:
            session = connect_to_sap()
            print(f"Linha atual: {int_num_lordem}.")
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
             ) = consulta_os(ordem, contrato, unadm)
            # Consulta Status da Ordem
            if status_sistema == fechada:
                print(f"Status do Sistema: {status_sistema}")
            else:
                print(f"OS: {ordem} aberta.")
                int_num_lordem += 1
                # Incremento + de Ordem.
                ordem = pendentes_list[int_num_lordem]
                continue

            if status_usuario == valorada:
                print(f"OS: {ordem} já valorada.")
                ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
                ja_valorado.valorada("SIM")
                int_num_lordem += 1
                # Incremento + de Ordem.
                ordem = pendentes_list[int_num_lordem]

            else:
                # Ação no SAP
                contrato_gbitaquera = "4600042888"
                contrato_novasp = "4600041302"
                if contrato == contrato_gbitaquera:
                    gbitaquera(ordem)
                elif contrato == contrato_novasp:
                    novasp(ordem)
                else:
                    print("Contrato não informado, encerrando.")
                    sys.exit()

                print("****Processo de Serviços Executados****")
                try:
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
                    )
                # pylint: disable=E1101
                except pywintypes.com_error:
                    print(f"Ordem: {ordem} em medição definitiva.")
                    ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
                    ja_valorado.valorada(obs="Definitiva")
                    int_num_lordem += 1
                    # Incremento + de Ordem.
                    ordem = pendentes_list[int_num_lordem]
                    continue

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
                        ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
                        ja_valorado.valorada(obs="SIM")
                        int_num_lordem += 1
                        # Incremento + de Ordem.
                        ordem = pendentes_list[int_num_lordem]
                        continue

                # pylint: disable=E1101
                except pywintypes.com_error:
                    print("OS Livre para valorar.")
                    session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS").select()
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                        + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
                    )

                # TSE e Aba Itens de preço
                (tse_proibida,
                    identificador,
                    list_chave_rb_despesa,
                    list_chave_unitario,
                    chave_rb_investimento,
                    chave_unitario,
                    etapa_rem_base,
                    etapa_unitario) = precificador(tse, corte, relig, posicao_rede, profundidade)

                # Se a TSE não estiver no escopo dar Val, vai pular pra próxima OS.
                if tse_proibida is not None:
                    ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
                    ja_valorado.valorada(obs="Num Pode")
                    int_num_lordem += 1
                    ordem = pendentes_list[int_num_lordem]
                    continue
                else:
                    # Aba Materiais
                    if chave_rb_investimento:
                        materiais(int_num_lordem,
                                  hidro,
                                  operacao,
                                  chave_rb_investimento,
                                  diametro_ramal,
                                  diametro_rede,
                                  contrato)

                    if list_chave_rb_despesa:
                        for chave_rb_despesa in list_chave_rb_despesa:
                            materiais(int_num_lordem,
                                      hidro,
                                      operacao,
                                      chave_rb_despesa,
                                      diametro_ramal,
                                      diametro_rede,
                                      contrato)

                    if list_chave_unitario:
                        for chave_unitario in list_chave_unitario:
                            materiais(int_num_lordem,
                                      hidro,
                                      operacao,
                                      chave_unitario,
                                      diametro_ramal,
                                      diametro_rede,
                                      contrato)
                    # Fim dos materiais
                    sys.exit()
                    # Salvar Ordem
                    qtd_ordem = salvar(ordem, qtd_ordem, contrato, unadm)
                    # Fim do contador de valoração.
                    cronometro_val(start_time, ordem)
                    # Incremento + de Ordem.
                    int_num_lordem += 1
                    ordem = pendentes_list[int_num_lordem]
                    print(f"Quantidade de ordens valoradas: {qtd_ordem}.")

        # pylint: disable=E1101
        except pywintypes.com_error as erro:
            # Baixa e abre novo arquivo do SAP
            print(f"Erro: {erro}")
            down_sap()
            print("Reiniciando programa")

    validador = True
    return ordem, int_num_lordem, validador
