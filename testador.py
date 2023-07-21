# testador.py
'''Área de testes.'''
from sap_connection import connect_to_sap
session = connect_to_sap()

preco = session.findById(
    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
    + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
preco.GetCellValue(0, "NUMERO_EXT")
num_precos_linhas = preco.RowCount
n_preco = 0  # índice para itens de preço
contador_pg = 0
# Definir o número de linhas visíveis
num_linhas_visiveis = 48
grid = session.findById(
    "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")

# Rolar para baixo até o final do grid


for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):

    sap_preco = preco.GetCellValue(
        n_preco, "NUMERO_EXT")
    item_preco = preco.GetCellValue(
        n_preco, "ITEM")
    n_etapa = preco.GetCellValue(
        n_preco, "ETAPA")

    # 660 é módulo despesa.

    if sap_preco == '456041' and item_preco == '660' \
            and n_etapa == '0040':
        preco.modifyCell(n_preco, "QUANT", "1")
        preco.setCurrentCell(n_preco, "QUANT")
        contador_pg += 1
        print("oi")
    else:
        print(f" olha ai:, n:{n_preco} {sap_preco, item_preco, n_etapa}")
        if n_preco >= num_linhas_visiveis:
            preco.currentCellRow = num_linhas_visiveis * 2
        if n_preco > num_linhas_visiveis * 2:
            preco.currentCellRow = num_linhas_visiveis * 4
