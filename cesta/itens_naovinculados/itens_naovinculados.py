#m_itens_naovinculados.py
import sys
from SAPConnection import Connect_to_SAP
session = Connect_to_SAP() 

class RemBaseReposicao:
    def FechamentoeReabertura():
        print("Iniciando processo de Modalide - REM BASE - MOD DESP FECH E REAB LIG - CÓDIGO: 327041 - PE ou 327051 - SM")
        modalidade = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        modalidade.GetCellValue(0, "NUMERO_EXT")
        if modalidade is not None:
            num_modalidade_linhas = modalidade.RowCount
            print(f"Quantidade linhas de modalidade: {num_modalidade_linhas}.")
            n_modalidade = 0
            for n_modalidade in range(num_modalidade_linhas):
                SAP_modalidade = modalidade.GetCellValue(n_modalidade, "NUMERO_EXT")
                if SAP_modalidade == str(327041) or SAP_modalidade == str(327051):
                    modalidade.modifyCell(n_modalidade, "MEDICAO", True)
                    modalidade.SetCurrentCell(n_modalidade, "MEDICAO")
                    modalidade.pressf4()
                    print(f"Selecionado modalidade: {SAP_modalidade}")
                    break
        #session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").selectColumn("ETAPA")
        
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").pressF4()
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").setCurrentCell(-1, "TSE")
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").selectColumn("TSE")
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").doubleClickCurrentCell()
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").setCurrentCell(-1, "NUMERO_EXT")
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").selectColumn("NUMERO_EXT") # Codigo da modalidade
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").doubleClickCurrentCell()
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").setCurrentCell(-1, "NUMERO_EXT")
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").selectColumn("NUMERO_EXT")
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").doubleClickCurrentCell()
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").setCurrentCell(-1, "MEDICAO")
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").selectColumn("MEDICAO") # Area de marcar.
        # session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell").doubleClickCurrentCell()
        
        
        
        
    def ManutLigacaoEsgoto(n_linha):
        print("Iniciando processo de Modalidade - REM BASE - MOD DESP MANUT LIG ESG - CÓDIGO: 7042")
        modalidade = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        modalidade.GetCellValue(0, n_linha)
        pass
    
    def ManutRedeEsgoto(n_linha):
        print("Iniciando processo de Modalidade - REM BASE - DESP MANUT REDE ESG - CÓDIGO: 7043")
        modalidade = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        modalidade.GetCellValue(0, n_linha)
        pass
    
    def ManutRedeAgua(n_linha):
        print("Iniciando processo de Modalidade - REM BASE - MOD DESP RP MAN RD AGUA - CÓDIGO: 7045")
        modalidade = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        modalidade.GetCellValue(0, n_linha)
        pass
    
    def ManutLigacaoAgua(n_linha):
        print("Iniciando processo de Modalidade - REM BASE - MOD DESP RP MAN LIG AGUA - CÓDIGO: 7046")
        modalidade = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        modalidade.GetCellValue(0, n_linha)
        pass
        
    def InvTrocaRamalAgua(n_linha):
        print("Iniciando processo de Modalidade - REM BASE - MOD INVEST TR LIG AGUA - CÓDIGO: 7050")
        modalidade = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
        modalidade.GetCellValue(0, n_linha)
        pass
    
    
    
RemBaseReposicao.FechamentoeReabertura()