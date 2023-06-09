#cavalete.py
import sys
from SAPConnection import Connect_to_SAP
from Excel_tbs import  load_worksheets
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

class Cavalete:
     
    @staticmethod
    def TrocaPeCvPrev(): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo Pai de Troca de Pé de Cavalete Preventiva - TSE 153000")
        servico_temp = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, SAP_tse in enumerate(range(0, num_tse_linhas)):
            SAP_tse = servico_temp.GetCellValue(n_tse, "TSE")
            
            if SAP_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                tse_temp_reposicao.append(SAP_tse)
                continue
            
            elif SAP_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n") # Cesta
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                servico_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
                continue
             
    @staticmethod    
    def TrocaCvKit(): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo Pai de Troca Cavalete KIT - TSE 149000")
        servico_temp = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, SAP_tse in enumerate(range(0, num_tse_linhas)):
            SAP_tse = servico_temp.GetCellValue(n_tse, "TSE")
            
            if SAP_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                tse_temp_reposicao.append(SAP_tse)
                continue
            
            elif SAP_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n") # Cesta
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                servico_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
                continue
                    
    @staticmethod    
    def RegularizarCv(): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo Pai de Regularizar Cavalete - TSE 142000")
        servico_temp = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, SAP_tse in enumerate(range(0, num_tse_linhas)):
            SAP_tse = servico_temp.GetCellValue(n_tse, "TSE")
            
            if SAP_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                tse_temp_reposicao.append(SAP_tse)
                continue
            
            elif SAP_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n") # Cesta
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                servico_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
                continue

    @staticmethod    
    def TrocaCvporUMA(): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo Pai de Troca de Cavalete por UMA - TSE 148000")
        servico_temp = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, SAP_tse in enumerate(range(0, num_tse_linhas)):
            SAP_tse = servico_temp.GetCellValue(n_tse, "TSE")
            
            if SAP_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                tse_temp_reposicao.append(SAP_tse)
                continue
            
            elif SAP_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n") # Cesta
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                servico_temp.append(SAP_tse) # Coloca a tse existente na lista temporária
                continue

