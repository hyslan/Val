import sys
from SAPConnection_n2 import Connect_to_SAP_n2
session_n2 = Connect_to_SAP_n2()


print("Iniciando processo de Modalide - REM BASE - MOD DESP FECH E REAB LIG - CÓDIGO: 327041 - PE ou 327051 - SM")
modalidade = session_n2.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV/ssubSUB_TAB:ZSBMM_VALORACAOINV:9035/cntlCC_ITNS_NVINCRB/shellcont/shell")
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
            