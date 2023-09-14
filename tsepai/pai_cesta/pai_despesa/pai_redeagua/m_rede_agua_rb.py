# m_rede_agua_rb.py
'''Módulo família Rede de Água (Ramal) - remuneração base '''
# Bibliotecas

from sap_connection import connect_to_sap
from excel_tbs import load_worksheets

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


class RedeAgua:
    '''Familia Rede de água'''
    MODALIDADE = "rede_agua"
    OBS = None

    @staticmethod
    def reparo_de_rede_de_agua():
        '''REPARO DE REDE DE ÁGUA'''
        etapa_reposicao = []
        tse_proibida = RedeAgua.OBS
        identificador = RedeAgua.MODALIDADE
        print("Iniciando processo Pai de REPARO DE REDE DE ÁGUA - TSE 332000")
        servico_temp = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                continue

            elif sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao

    @staticmethod
    def aperto_gaxeta_valvula():
        '''APERTO DE GAXETA VALVULA REDE DE AGUA'''
        etapa_reposicao = []
        tse_proibida = RedeAgua.OBS
        identificador = "gaxeta"
        print(
            "Iniciando processo Pai de APERTO DE GAXETA VALVULA REDE DE AGUA - TSE 328000")
        servico_temp = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                continue

            elif sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao
