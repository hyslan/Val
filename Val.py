'''Sistema Val: programa de valoração automática não assistida, Author: Hyslan Silva Cruz'''
#main.py
#Bibliotecas
import time
import datetime
from tqdm import tqdm # Barra de progresso
from Excel_tbs import load_worksheets
from unitarios import dicionario
from SAPConnection import Connect_to_SAP
from ConfereOS import ConsultaOS
from ZSBMM216 import novasp
from ServicosExecutados import VerificaTSE


# Função Principal
def main():
    '''Sistema principal da Val e inicializador do programa'''
    hora_parada = datetime.time(21, 50) # Ponto de parada às 21h50min
    hora_retomada = datetime.time(6, 0) # Ponto de retomada às 6h
    while True:
        hora_atual = datetime.datetime.now().time() # Obtém a hora atual
        print(f"Hora atual: {hora_atual}")
        # Conexão SAP
        session = Connect_to_SAP()
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
        tb_tse_UN,
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
        except:
            print("Entrada inválida. Digite um número inteiro válido.")
            print("Reiniciando o programa...")
            main()
                        
        # Variáveis de Status da Ordem
        Valorada = "EXEC VALO" or "NEXE VALO"
        Executado = "EXEC" or "NEXE"
        Fechada = "LIB" 
        print(f"Ordem selecionada: {ordem} \n Linha: {int_num_lordem}")
        qtdOrdem = 0 # Contador de ordens pagas.
        for num_lordem in tqdm(range(limite_execucoes), ncols=100): # Loop para pagar as ordens da planilha do Excel
            MaterialObs = planilha.cell(row = int_num_lordem, column = 3)
            SelecaoCarimbo = planilha.cell(row = int_num_lordem, column = 2)
            OrdemObs = planilha.cell(row = int_num_lordem, column = 4)
            print(f"Linha atual: {int_num_lordem}.")  
            start_time = time.time() # Contador de tempo para valorar.
            print(f"Ordem atual: {ordem}")
            print("Verificando Status da Ordem.")
            ConsultaOS(ordem) # Função consulta de Ordem.
            print("Iniciando Consulta.")
            StatusSistema, StatusUsuario, Corte, Relig, PosicaoRede, Profundidade = ConsultaOS(ordem)
            # Consulta Status da Ordem
            if StatusSistema == Fechada:
                print(f"Status do Sistema: {StatusSistema}")
            else:
                print(f"OS: {ordem} aberta.")
                SelecaoCarimbo = planilha.cell(row = int_num_lordem, column = 2)
                SelecaoCarimbo.value = "OS ABERTA"
                lista.save('lista.xlsx') # salva Planilha 
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value # Incremento + de Ordem.
                continue
                    
            if StatusUsuario == Valorada:
                print(f"OS: {ordem} já valorada.")
                SelecaoCarimbo = planilha.cell(row = int_num_lordem, column = 2)
                SelecaoCarimbo.value = "VALORADA ANTERIORMENTE"
                lista.save('lista.xlsx') # salva Planilha 
                int_num_lordem += 1
                ordem = planilha.cell(row=int_num_lordem, column=1).value # Incremento + de Ordem. 
                continue
            else:
            # Ação no SAP
                print("Iniciando valoração.")
                SessaoBotoes = session.findById("wnd[0]") # Sessão 0
                novasp() # Contrato NOVASP
                SAP_ordem = session.findById("wnd[0]/usr/ctxtP_ORDEM") # Campo ordem
                SAP_ordem.Text = (ordem)
                SessaoBotoes.SendVkey(8)# Aperta botão F8 
                print("*****************************************Processo de Serviços Executados************************************")
                try:
                    tse = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
                except:
                    print(f"Ordem: {ordem} em medição definitiva ou com erro.")
                    OrdemObs = planilha.cell(row = int_num_lordem, column = 4)
                    OrdemObs.value = "MEDIÇÃO DEFINITIVA OU COM ERRO."
                    int_num_lordem += 1
                    ordem = planilha.cell(row=int_num_lordem, column=1).value # Incremento + de Ordem.
                    continue
                
                tse.GetCellValue(0, "TSE") # Saber qual TSE é
                if tse is not None:
                    tse_temp, reposicao = VerificaTSE(tse)  
                    # Aba Itens de preço 
                    session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI").select()
                    print("*****************************************Processo de Precificação************************************")
                    for Etapa in tse_temp:
                        print(f"Tse temporária selecionada para pagar: {Etapa}")
                        if Etapa in tb_tse_UN: # Verifica se está no Conjunto Unitários
                            print(f"{Etapa} é unitário!")
                            dicionario.Unitario(Etapa, Corte, Relig)          
                        else:
                            print(f"{Etapa} não é unitário")
                            
                        # Aba Materiais
                        session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM").select()
                        print("*****************************************Processo de Materiais************************************")
                        tb_materiais = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM/ssubSUB_TAB:ZSBMM_VALORACAOINV:9030/cntlCC_MATERIAIS/shellcont/shell")
                        try:
                            SAP_Material = tb_materiais.GetCellValue(0, "MATERIAL") # Valor da célula
                            if SAP_Material is not None:  # Se existem materiais, irá contar
                                num_material_linhas = tb_materiais.RowCount # Conta as Rows
                                print(f"Qtd de linhas de materiais: {num_material_linhas}")
                                n_material = 0 # Número da Row do Grid Materiais do SAP
                                for n_material in range(num_material_linhas): # Loop do Grid Materiais, verifica se tem material da contratada e marca
                                    SAP_Material = tb_materiais.GetCellValue(n_material, "MATERIAL") # Pega valor da célula 0
                                    if SAP_Material in tb_contratada: # Verifica se está na lista tb_contratada
                                        tb_materiais.modifyCheckbox(n_material, "CONTRATADA", True) # Marca Contratada
                                        print(f"Linha do material: {n_material}, Material: {SAP_Material}")
                                        continue
                        except:
                            MaterialObs = planilha.cell(row = int_num_lordem, column = 3)
                            MaterialObs.value =  "Sem Material Vinculado"
                            print("Sem material vinculado.")
                                                                                     
                        # Fim dos materiais
                        SessaoBotoes.sendVKey(11) # Salvar Ordem
                        session.findById("wnd[1]/usr/btnBUTTON_1").press()
                        print("Salvando valoração!")
                        # Verificar se Salvou
                        print("Verificando se Ordem foi valorada.")
                        ConsultaOS(ordem)
                        print("Iniciando processo de verificação.")
                        StatusSistema, StatusUsuario, Corte, Relig, PosicaoRede, Profundidade = ConsultaOS(ordem)
                        if StatusUsuario == "EXEC VALO":
                            print(f"Status da Ordem: {StatusSistema}, {StatusUsuario}")
                            print("Foi Salvo com sucesso!")
                            SelecaoCarimbo = planilha.cell(row = int_num_lordem, column = 2)
                            SelecaoCarimbo.value = "VALORADA"
                            lista.save('lista.xlsx') # salva Planilha 
                            qtdOrdem += 1
                        else:
                            print(f"Ordem: {ordem} não foi salva.")
                            SelecaoCarimbo = planilha.cell(row = int_num_lordem, column = 2)
                            SelecaoCarimbo.value = "NÃO FOI SALVO"
                            lista.save('lista.xlsx') # salva Planilha 
                            
                        
                        # Fim do contador de valoração.
                        end_time = time.time()
                        execution_time = end_time - start_time # Tempo de execução.
                        print(f"Tempo gasto para valorar a Ordem: {ordem}, foi de: {execution_time} segundos.")
                        print(f"*****************************************Fim da Valoração da Ordem: {ordem} *****************************************")
                        int_num_lordem += 1
                        ordem = planilha.cell(row=int_num_lordem, column=1).value # Incremento + de Ordem.
                        print(f"Quantidade de ordens valoradas: {qtdOrdem}.")
                        lista.save('lista.xlsx') # salva Planilha
                                 
        if hora_atual >= hora_parada:
            print("A Val foi descansar.")
            print("- Val: até amanhã.")
            while True: # Loop da parada
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

