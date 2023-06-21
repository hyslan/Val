# core.py
'''Coração da Val.'''
import time
import pywintypes
from tqdm import tqdm
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from confere_os import consulta_os
from transact_zsbmm216 import novasp
from pagador import precificador
from almoxarifado import materiais
from salvacao import salvar
from temporizador import cronometro_val
session = connect_to_sap()
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
    _,
    *_,
) = load_worksheets()


def val():
    '''Sistema Val.'''
    validador = False
    input("- Val: Pressione Enter para iniciar...")
    limite_execucoes = planilha.max_row
    print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
    try:
        num_lordem = input("Insira o número da linha aqui: ")
        int_num_lordem = int(num_lordem)
        ordem = planilha.cell(row=int_num_lordem, column=1).value
    except TypeError:
        print("Entrada inválida. Digite um número inteiro válido.")
        print("Reiniciando o programa...")
        val()
    # Variáveis de Status da Ordem
    valorada = "EXEC VALO" or "NEXE VALO"
    fechada = "LIB"
    print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")
    qtd_ordem = 0  # Contador de ordens pagas.
    # Loop para pagar as ordens da planilha do Excel
    for num_lordem in tqdm(range(int_num_lordem, limite_execucoes), ncols=100):
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        ordem_obs = planilha.cell(row=int_num_lordem, column=4)
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
            _,
            _,
            hidro,
            operacao
         ) = consulta_os(ordem)
        # Consulta Status da Ordem
        if status_sistema == fechada:
            print(f"Status do Sistema: {status_sistema}")
        else:
            print(f"OS: {ordem} aberta.")
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "OS ABERTA"
            # salva Planilha
            lista.save('lista.xlsx')
            int_num_lordem += 1
            # Incremento + de Ordem.
            ordem = planilha.cell(row=int_num_lordem, column=1).value
            continue
        if status_usuario == valorada:
            print(f"OS: {ordem} já valorada.")
            selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
            selecao_carimbo.value = "VALORADA ANTERIORMENTE"
            lista.save('lista.xlsx')  # salva Planilha
            int_num_lordem += 1
            # Incremento + de Ordem.
            ordem = planilha.cell(row=int_num_lordem, column=1).value
            continue
        else:
            # Ação no SAP
            novasp(ordem)
            print("****Processo de Serviços Executados****")
            try:
                tse = session.findById(
                    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
                )
            # pylint: disable=E1101
            except pywintypes.com_error:
                print(f"Ordem: {ordem} em medição definitiva ou com erro.")
                ordem_obs = planilha.cell(row=int_num_lordem, column=4)
                ordem_obs.value = "MEDIÇÃO DEFINITIVA OU COM ERRO."
                lista.save('lista.xlsx')
                # Incremento de Ordem.
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value
                continue
            # TSE e Aba Itens de preço
            tse_proibida, identificador = precificador(tse, corte, relig)
            if tse_proibida is not None:
                selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
                selecao_carimbo.value = "Forbidden TSE"
                lista.save('lista.xlsx')  # salva Planilha
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value
                continue
            else:
                # Aba Materiais
                materiais(int_num_lordem, hidro, operacao, identificador)
                # Fim dos materiais
                # Salvar Ordem
                qtd_ordem = salvar(ordem, int_num_lordem, qtd_ordem)
                # Fim do contador de valoração.
                cronometro_val(start_time, ordem)
                # Incremento + de Ordem.
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value
                print(f"Quantidade de ordens valoradas: {qtd_ordem}.")
                lista.save('lista.xlsx')  # salva Planilha
    validador = True
    return ordem, int_num_lordem, validador
