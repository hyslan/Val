#m_supressao.py
'''Módulo Família Supressão Unitário.'''
from sap_connection import connect_to_sap
from excel_tbs import  load_worksheets
session = connect_to_sap()

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

class Supressao:
    '''Classe Pai de Supressão Unitário.'''
    @staticmethod
    def suprimir_ligacao_de_agua():
        '''Módulo Supressão Unitário.'''
        tse_proibida = None
        etapa_reposicao = None
        identificador = "supressao"
        print("Iniciando processo Pai de SUPRIMIR LIG ÁGUA - TSE: 405000 e TSE 414000")
        servico_temp = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                                        + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")
            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "s")
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao = etapa
                print(f"Tem reposição TSE: {sap_tse}")
                continue
            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n") # Cesta
                servico_temp.modifyCell(n_tse, "CODIGO", "3") # Pertence ao serviço principal
                servico_temp.append(sap_tse) # Coloca a tse existente na lista temporária
                continue
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao
    