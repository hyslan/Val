# m_rede_esgoto_rb.py
'''Módulo família Rede de Esgoto - remuneração base '''
# Bibliotecas
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets


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
    tb_contratada_gb,
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


class RedeEsgoto:
    '''Familia Rede de Esgoto'''
    MODALIDADE = "rede_esgoto"
    OBS = None

    @staticmethod
    def reparo_de_rede_de_esgoto():
        '''REPARO DE Rede DE ESGOTO - RB'''
        session = connect_to_sap()
        etapa_reposicao = []
        tse_proibida = RedeEsgoto.OBS
        identificador = RedeEsgoto.MODALIDADE
        print("Iniciando processo Pai de REPARO DE REDE DE ESGOTO - TSE 580000")
        servico_temp = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
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
