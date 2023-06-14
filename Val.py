'''Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz'''
#main.py
#Bibliotecas
#pylint: disable=W0611
import sys
import time
import datetime
import pywintypes
from tqdm import tqdm # Barra de progresso
from excel_tbs import load_worksheets
from unitarios import dicionario
from sap_connection import connect_to_sap
from confere_os import consulta_os
from zsbmm2216 import novasp
from servicos_executados import verifica_tse


# Função Principal
def main():
    '''Sistema principal da Val e inicializador do programa'''
    hora_parada = datetime.time(21, 50) # Ponto de parada às 21h50min
    hora_retomada = datetime.time(6, 0) # Ponto de retomada às 6h
    while True:
        hora_atual = datetime.datetime.now().time() # Obtém a hora atual
        print(f"Hora atual: {hora_atual}")
        # Conexão SAP
        session = connect_to_sap()
        #Área do Excel, definição das variáveis
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
        tb_contratada,
        tb_tse_un,
        *_,
        ) = load_worksheets()

        #Início do Sistema
        print(" - Val:")
        input("Pressione Enter para iniciar...")
        limite_execucoes = planilha.max_row
        print(f"Quantidade de ordens incluídas na lista: {limite_execucoes}")
        try: #Fazer uma função separada
            num_lordem = input("Insira o número da linha aqui:")
            int_num_lordem = int(num_lordem)
            ordem = planilha.cell(row=int_num_lordem, column=1).value
        except TypeError:
            print("Entrada inválida. Digite um número inteiro válido.")
            print("Reiniciando o programa...")
            main()
        # Variáveis de Status da Ordem
        valorada = "EXEC VALO" or "NEXE VALO"
        fechada = "LIB"
        print(f"Ordem selecionada: {ordem} , Linha: {int_num_lordem}")
        qtd_ordem = 0 # Contador de ordens pagas.
        # Loop para pagar as ordens da planilha do Excel
        for num_lordem in tqdm(range(int_num_lordem, limite_execucoes), ncols=100):
            material_obs = planilha.cell(row = int_num_lordem, column = 3)
            selecao_carimbo = planilha.cell(row = int_num_lordem, column = 2)
            ordem_obs = planilha.cell(row = int_num_lordem, column = 4)
            print(f"Linha atual: {int_num_lordem}.")
            start_time = time.time() # Contador de tempo para valorar.
            print(f"Ordem atual: {ordem}")
            print("Verificando Status da Ordem.")
            consulta_os(ordem) # Função consulta de Ordem.
            print("Iniciando Consulta.")
            status_sistema, status_usuario, corte, relig, _, _, hidro_instalado = consulta_os(ordem)
            # Consulta Status da Ordem
            if status_sistema == fechada:
                print(f"Status do Sistema: {status_sistema}")
            else:
                print(f"OS: {ordem} aberta.")
                selecao_carimbo = planilha.cell(row = int_num_lordem, column = 2)
                selecao_carimbo.value = "OS ABERTA"
                # salva Planilha
                lista.save('lista.xlsx')
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value # Incremento + de Ordem.
                continue
            if status_usuario == valorada:
                print(f"OS: {ordem} já valorada.")
                selecao_carimbo = planilha.cell(row = int_num_lordem, column = 2)
                selecao_carimbo.value = "VALORADA ANTERIORMENTE"
                lista.save('lista.xlsx') # salva Planilha
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value # Incremento + de Ordem.
                continue
            else:
            # Ação no SAP
                print("Iniciando valoração.")
                sessao_botoes = session.findById("wnd[0]") # Sessão 0
                novasp() # Contrato NOVASP
                sap_ordem = session.findById("wnd[0]/usr/ctxtP_ORDEM") # Campo ordem
                sap_ordem.Text = ordem
                sessao_botoes.SendVkey(8) # Aperta botão F8
                print("****Processo de Serviços Executados****")
                try:
                    tse = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                       + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell"
                       )
                #pylint: disable=E1101
                except pywintypes.com_error:
                    print(f"Ordem: {ordem} em medição definitiva ou com erro.")
                    ordem_obs = planilha.cell(row = int_num_lordem, column = 4)
                    ordem_obs.value = "MEDIÇÃO DEFINITIVA OU COM ERRO."
                    # Incremento + de Ordem.
                    int_num_lordem += 1
                    ordem = planilha.cell(row=int_num_lordem, column=1).value
                    continue
                tse.GetCellValue(0, "TSE") # Saber qual TSE é
                if tse is not None:
                    tse_temp, reposicao, num_tse_linhas = verifica_tse(tse)
                    print(f"TSE: {tse_temp}, Reposição inclusa ou não: {reposicao}")
                    # Aba Itens de preço
                    session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI").select()
                    print("****Processo de Precificação****")
                    for etapa in tse_temp:
                        print(f"TSE selecionada para pagar: {etapa}")
                        if etapa in tb_tse_un: # Verifica se está no Conjunto Unitários
                            print(f"{etapa} é unitário!")
                            #pylint: disable=E1121
                            dicionario.unitario(etapa, corte, relig, reposicao, num_tse_linhas)
                        else:
                            print(f"{etapa} não é unitário!")
                    # Aba Materiais
                    session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM").select()
                    print("****Processo de Materiais****")
                    tb_materiais = session.findById(
                        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM/ssubSUB_TAB:"
                    + "ZSBMM_VALORACAOINV:9030/cntlCC_MATERIAIS/shellcont/shell"
                                                    )
                    try:
                        # Valor da célula
                        sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
                        if sap_material is not None:  # Se existem materiais, irá contar
                            num_material_linhas = tb_materiais.RowCount # Conta as Rows
                            print(f"Qtd de linhas de materiais: {num_material_linhas}")
                            # Número da Row do Grid Materiais do SAP
                            n_material = 0
                            ultima_linha_material = num_material_linhas
                            hidro_y = 'Y'
                            # Hidrômetro atual.
                            hidro_instalado = hidro_instalado.upper()
                            # Mata-burro pra hidro.
                            if hidro_instalado.startswith(hidro_y):
                                cod_hidro_instalado = '50000108'
                            else:
                                cod_hidro_instalado = '50000530'
                            # Variável para controlar se o hidrômetro já foi adicionado
                            hidro_adicionado = False
                            # Loop do Grid Materiais.
                            for n_material in range(num_material_linhas):
                                # Pega valor da célula 0
                                sap_material = tb_materiais.GetCellValue(n_material, "MATERIAL")
                                sap_etapa_material = tb_materiais.GetCellValue(
                                    n_material, "ETAPA")
                                # Verifica se está na lista tb_contratada
                                if sap_material in tb_contratada:
                                    # Marca Contratada
                                    tb_materiais.modifyCheckbox(
                                        n_material, "CONTRATADA", True)
                                    print(f"Linha do material: {n_material}, "
                                        + f"Material: {sap_material}")
                                    continue
                                if sap_material == '50000328':
                                    tb_materiais.modifyCheckbox(
                                        n_material, "ELIMINADO", True
                                    )
                                    tb_materiais.InsertRows(str(ultima_linha_material))
                                    tb_materiais.modifyCell(
                                        ultima_linha_material, "ETAPA", sap_etapa_material
                                        )
                                    tb_materiais.modifyCell(
                                        ultima_linha_material, "MATERIAL", "50000263"
                                        )
                                    tb_materiais.modifyCell(
                                        ultima_linha_material, "QUANT", "1"
                                        )
                                    tb_materiais.setCurrentCell(
                                        ultima_linha_material, "QUANT"
                                        )
                                    ultima_linha_material = ultima_linha_material + 1

                                if hidro_instalado is not None:
                                    if hidro_adicionado is True:
                                        print("Hidro já adicionado.")
                                    else:
                                        if sap_material == cod_hidro_instalado:
                                            print(f"Hidro foi incluso corretamente: {cod_hidro_instalado}")
                                            # Hidrômetro foi adicionado
                                            hidro_adicionado = True

                                        elif sap_material == '50000108' and sap_material != cod_hidro_instalado:
                                            print(f"Hidro inserido incorretamente, incluindo o informado: {cod_hidro_instalado}")
                                            tb_materiais.modifyCheckbox(n_material, "ELIMINADO", True)
                                            tb_materiais.InsertRows(str(ultima_linha_material))
                                            tb_materiais.modifyCell(ultima_linha_material, "ETAPA", sap_etapa_material)
                                            tb_materiais.modifyCell(ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                                            tb_materiais.modifyCell(ultima_linha_material, "QUANT", "1")
                                            tb_materiais.setCurrentCell(ultima_linha_material, "QUANT")
                                            ultima_linha_material = ultima_linha_material + 1
                                            hidro_adicionado = True  # Hidrômetro foi adicionado

                                        elif sap_material == '50000530' and sap_material != cod_hidro_instalado:
                                            print(f"Hidro inserido incorretamente, incluindo o informado: {cod_hidro_instalado}")
                                            tb_materiais.modifyCheckbox(n_material, "ELIMINADO", True)
                                            tb_materiais.InsertRows(str(ultima_linha_material))
                                            tb_materiais.modifyCell(ultima_linha_material, "ETAPA", sap_etapa_material)
                                            tb_materiais.modifyCell(ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                                            tb_materiais.modifyCell(ultima_linha_material, "QUANT", "1")
                                            tb_materiais.setCurrentCell(ultima_linha_material, "QUANT")
                                            ultima_linha_material = ultima_linha_material + 1
                                            hidro_adicionado = True  # Hidrômetro foi adicionado

                            if hidro_instalado is not None and hidro_adicionado is False:
                                print(f"Não foi inserido hidro, incluindo o informado: {cod_hidro_instalado}")
                                tb_materiais.InsertRows(str(ultima_linha_material))
                                tb_materiais.modifyCell(ultima_linha_material, "ETAPA", sap_etapa_material)
                                tb_materiais.modifyCell(ultima_linha_material, "MATERIAL", cod_hidro_instalado)
                                tb_materiais.modifyCell(ultima_linha_material, "QUANT", "1")
                                tb_materiais.setCurrentCell(ultima_linha_material, "QUANT")
                                ultima_linha_material = ultima_linha_material + 1
                                hidro_adicionado = True # Hidrômetro foi adicionado
                    #pylint: disable=E1101
                    except pywintypes.com_error:
                        material_obs = planilha.cell(row = int_num_lordem, column = 3)
                        material_obs.value =  "Sem Material Vinculado"
                        print("Sem material vinculado.")
                        #inserir materiais aqui.
                    # Fim dos materiais
                    sys.exit()
                    sessao_botoes.sendVKey(11) # Salvar Ordem
                    session.findById("wnd[1]/usr/btnBUTTON_1").press()
                    print("Salvando valoração!")
                    # Verificar se Salvou
                    print("Verificando se Ordem foi valorada.")
                    consulta_os(ordem)
                    print("Iniciando processo de verificação.")
                    status_sistema, status_usuario, corte, relig, _, _, hidro_instalado = consulta_os(ordem)
                    if status_usuario == "EXEC VALO":
                        print(f"Status da Ordem: {status_sistema}, {status_usuario}")
                        print("Foi Salvo com sucesso!")
                        selecao_carimbo = planilha.cell(row = int_num_lordem, column = 2)
                        selecao_carimbo.value = "VALORADA"
                        lista.save('lista.xlsx') # salva Planilha
                        qtd_ordem += 1
                    else:
                        print(f"Ordem: {ordem} não foi salva.")
                        selecao_carimbo = planilha.cell(row = int_num_lordem, column = 2)
                        selecao_carimbo.value = "NÃO FOI SALVO"
                        lista.save('lista.xlsx') # salva Planilha
                    # Fim do contador de valoração.
                    end_time = time.time()
                    execution_time = end_time - start_time # Tempo de execução.
                    print(f"Tempo gasto para valorar a Ordem: {ordem}, "
                            + f"foi de: {execution_time} segundos.")
                    print(f"****Fim da Valoração da Ordem: {ordem} ****")
                    # Incremento + de Ordem.
                    int_num_lordem += 1
                    ordem = planilha.cell(row=int_num_lordem, column=1).value
                    print(f"Quantidade de ordens valoradas: {qtd_ordem}.")
                    lista.save('lista.xlsx') # salva Planilha
        # Loop de Parada
        if hora_atual >= hora_parada:
            print("A Val foi descansar.")
            print("- Val: até amanhã.")
            while True:
                hora_atual = datetime.datetime.now().time()
                print(f"Hora atual na parada: {hora_atual}")
                time.sleep(60)
                if hora_atual >= hora_retomada:
                    print("- Val: Bom dia!")
                    print("- Val: Vamos trabalhar!")
                    print(f"- Val: Retomando Ordem: {ordem} \n da linha: {int_num_lordem}")
                    break # Sai do Loop quando atingir a hora de retomada.

#-Main------------------------------------------------------------------
if __name__ == '__main__':
    main()
#-End-------------------------------------------------------------------
