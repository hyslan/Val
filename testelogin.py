'''Área de testes geral.'''
from lista_reposicao import dict_reposicao
from sap_connection import connect_to_sap
session = connect_to_sap()
servico = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                               + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
servico.DoubleClick(1, "ETAPA")
hidrometro_instalado = session.findById(
            "wnd[0]/usr/txtGS_AFVU-ZZHIDROMETRO_INSTALADO").Text
print(hidrometro_instalado)