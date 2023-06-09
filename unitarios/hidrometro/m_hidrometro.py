#hidrometro.py
from SAPConnection import Connect_to_SAP
session = Connect_to_SAP() 
class Hidrometro: 
    @staticmethod
    def THDPrev_456902(*args): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo de pagar TROCA DE HIDROMETRO PREVENTIVA - Código: 456902")
        preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount 
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0 # índice para itens de preço
            for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if SAP_preco == str(456902): # THD  ATE 10M3/H PREV 
                    preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de THD  ATE 10M3/H PREV - CODIGO: 456902")
                    break
                else:
                    print(f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")      
    @staticmethod    
    def HD_456022(*args): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
        print("Iniciando processo de pagar COLOCADO HIDROMETRO NA POSIÇÃO CORRETA  - Código: 456022")
        preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount 
            print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
            n_preco = 0 # índice para itens de preço
            for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if SAP_preco == str(456022): #  
                    preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    print("Pago 1 UN de DESINCL HD - CODIGO: 456022")
                    break
                else:
                    print(f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")
    @staticmethod
    def THD_456901(*args): # Deve Criar uma instância na main já com a instância da classe feita, exemplo: hidrometro_instancia.THDPrev()
            print("Iniciando processo de pagar TROCA DE HIDROMETRO CORRETIVO - Código: 456901")
            preco = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
            preco.GetCellValue(0, "NUMERO_EXT")
            if preco is not None:
                num_precos_linhas = preco.RowCount 
                print(f"Qtd linhas em itens de preço: {num_precos_linhas}")
                n_preco = 0 # índice para itens de preço
                for n_preco, SAP_preco in enumerate(range(0, num_precos_linhas)):
                    SAP_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                    if SAP_preco == str(456901): #  
                        preco.modifyCell(n_preco, "QUANT", "1") # Marca pagar na TSE
                        preco.setCurrentCell(n_preco, "QUANT")
                        preco.pressEnter()
                        print("Pago 1 UN de THD  ATE 10M3/H CORR - CODIGO: 456902")
                        break
                    else:
                        print(f"Código de preço: {SAP_preco} , Linha: {n_preco} - Não Selecionado")
