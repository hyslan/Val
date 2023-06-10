#ServicosExecutados.py
import sys
from sap_connection import Connect_to_SAP
from excel_tbs import load_worksheets
from tsepai import pai_dicionario

session = Connect_to_SAP()
(
        lista,
        materiais,
        plan_tse,
        plan_precos,
        planilha,
        contratada,
        unitario,
        rb,
        naoexecutado,
        invest,
        tb_contratada,
        tb_tse_UN,
        tb_tse_rb,
        tb_tse_nexec,
        tb_tse_invest,
        tb_StSistema,
        tb_StUsuario,
        tb_tse_PertenceAoServicoPrincipal,
        tb_tse_ServicoNaoExistenoContrato,
        tb_tse_reposicao,
        tb_tse_Retrabalho,
        tb_tse_Asfalto,
        ) = load_worksheets()


def VerificaTSE(servico):
    print("Iniciando processo de verificação de TSE")
    servico = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
    n_tse = 0
    tse_temp = [] # Lista temporária para armazenar as tse
    num_tse_linhas = servico.RowCount 
    print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
    for n_tse, SAP_tse in enumerate(range(0, num_tse_linhas)):
        SAP_tse = servico.GetCellValue(n_tse, "TSE")
        
        if SAP_tse in tb_tse_UN: # Verifica se está no Conjunto Unitários
            servico.modifyCell(n_tse, "PAGAR", "s") # Marca pagar na TSE
            tse_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
            reposicao = pai_dicionario.PaiservicoUnitario(SAP_tse)
            
            continue
            
        elif SAP_tse in tb_tse_rb: # Caso Contrário, é RB - Despesa
            servico.modifyCell(n_tse, "PAGAR", "n") # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5") # Despesa
            tse_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
            pai_dicionario.PaiservicoCesta(SAP_tse)
            continue
          # tirar depois   
        elif SAP_tse in tb_tse_invest:# Caso Contrário, é RB - Investimento
            servico.modifyCell(n_tse, "PAGAR", "n") # Cesta
            servico.modifyCell(n_tse, "CODIGO", "6") # Investimento
            tse_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
            continue
            #tirar depois
        elif SAP_tse in tb_tse_nexec:
            servico.modifyCell(n_tse, "PAGAR", "n") # Cesta
            servico.modifyCell(n_tse, "CODIGO", "11") # Não Executado
            tse_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
            continue
            
        elif SAP_tse in tb_tse_PertenceAoServicoPrincipal:
            servico.modifyCell(n_tse, "PAGAR", "n") # 
            servico.modifyCell(n_tse, "CODIGO", "3") # Pertence ao Serviço Principal
            tse_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
            continue
        
        elif SAP_tse in tb_tse_Retrabalho:
            servico.modifyCell(n_tse, "PAGAR", "n") # 
            servico.modifyCell(n_tse, "CODIGO", "7") # Retrabalho
            tse_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
            continue
        
        else:
            print("TSE não encontrado na planilha do Excel.")
            #sys.exit()
    
    servico.pressEnter()
    
    # for valor_tse_temp in tse_temp:
    #     print(f"TSEs na lista temporária: {valor_tse_temp}")   
    return tse_temp, reposicao
