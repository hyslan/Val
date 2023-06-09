#religacao.py
#Incluir etapas nas intruções para classes superiores
import sys
from SAPConnection import Connect_to_SAP
session = Connect_to_SAP() 
class Religacao:
    @staticmethod    
    def Restabelecida(Corte, Relig):
        if Relig is not None:
            if Relig == 'CAVALETE':
                    print("Iniciando processo de pagar RELIG CV - Código: 456037")
                    preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                    preco.GetCellValue(0, "NUMERO_EXT")
                    if preco is not None:
                        num_precos_linhas = preco.RowCount 
                        print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
                        n_preco = 0 # índice para itens de preço
                        for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                            SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                            if SAP_preco == str(456037):   
                                preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                                preco.setCurrentCell(n_preco, "QUANT")
                                preco.pressEnter()
                                print("Pago 1 UN de RELIG CV - CODIGO: 456037")
                                break
                            
                            elif Relig == 'RAMAL':
                                print("Iniciando processo de pagar RELIG  RAMAL AG  S/REP - Código: 456039")
                                preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                                preco.GetCellValue(0, "NUMERO_EXT")
                                if preco is not None:
                                    num_precos_linhas = preco.RowCount 
                                    print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
                                    n_preco = 0 # índice para itens de preço
                                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                                        SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                                        if SAP_preco == str(456039):   
                                            preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                                            preco.setCurrentCell(n_preco, "QUANT")
                                            preco.pressEnter()
                                            print("Pago 1 UN de RELIG  RAMAL AG  S/REP - CODIGO: 456039")
                                            break
                                        else:
                                            print(f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")                        
                            
                            elif Relig == 'FERRULE':
                                print("Iniciando processo de pagar RELIG  TMD AG  S/REP - Código: 456040")
                                preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                                preco.GetCellValue(0, "NUMERO_EXT")
                                if preco is not None:
                                    num_precos_linhas = preco.RowCount 
                                    print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
                                    n_preco = 0 # índice para itens de preço
                                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                                        SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                                        if SAP_preco == str(456040):   
                                            preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                                            preco.setCurrentCell(n_preco, "QUANT")
                                            preco.pressEnter()
                                            print("Pago 1 UN de RELIG  TMD AG  S/REP - CODIGO: 456040")
                                            break
                                        else:
                                            print(f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")
            else:
                print("Religação não informada. \n Pagando como RELIG CV.")
                preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
                preco.GetCellValue(0, "NUMERO_EXT")
                if preco is not None:
                    num_precos_linhas = preco.RowCount 
                    print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
                    n_preco = 0 # índice para itens de preço
                    for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                        SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                        if SAP_preco == str(456037): #  
                            preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                            preco.setCurrentCell(n_preco, "QUANT")
                            preco.pressEnter()
                            print("Pago 1 UN de SUPR CV - CODIGO: 456037")
                