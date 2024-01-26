'''Módulo Família Poço Unitário.'''
from src.sap_connection import connect_to_sap
from src.excel_tbs import load_worksheets


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


class Poco:
    '''Classe Pai Poço.'''
    @staticmethod
    def det_descoberto_nivelado_reg_cx_parada():
        '''Módulo de  Descobrir, Trocar Caixa de Parada'''
        session = connect_to_sap()
        etapa_reposicao = []
        reposicao_tse_temp = []
        tse_proibida = None
        identificador = "cx_parada"
        print("Iniciando processo Pai de Det. Desc. Niv. Reg. Parada - "
              + "TSE: 322000")
        servico_temp = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            etapa = servico_temp.GetCellValue(n_tse, "ETAPA")

            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                servico_temp.append(sap_tse)
                continue

            if sap_tse in tb_tse_reposicao:
                if sap_tse in ('170301', '749000', '758500'):
                    servico_temp.modifyCell(n_tse, "PAGAR", "n")
                    servico_temp.modifyCell(n_tse, "CODIGO", "10")

                servico_temp.modifyCell(n_tse, "PAGAR", "s")
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

        return reposicao_tse_temp, tse_proibida, identificador, etapa_reposicao

    @staticmethod
    def nivelamento():
        '''Módulo de Nivelamento'''
        session = connect_to_sap()
        etapa_reposicao = []
        tse_proibida = None
        identificador = "poço"
        print("Iniciando processo Pai de Nivelamento")
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
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                continue

            elif sap_tse in tb_tse_ServicoNaoExistenoContrato:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao.append(etapa)

        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao
