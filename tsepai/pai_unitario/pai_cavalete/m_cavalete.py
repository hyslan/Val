# cavalete.py
'''Módulo Família Cavalete Unitário.'''
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


class Cavalete:
    '''Classe dos Cavaletes Unitários.'''
    TIPO = "cavalete"

    @staticmethod
    def troca_pe_cv_prev():
        '''Módulo Pai troca Pé de Cavalete.'''
        etapa_reposicao = None
        identificador = Cavalete.TIPO
        tse_proibida = None
        print("Iniciando processo Pai de Troca de Pé de Cavalete Preventiva - TSE 153000")
        servico_temp = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        tse_temp_reposicao = []
        num_tse_linhas = servico_temp.RowCount
        print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")

            if sap_tse in tb_tse_reposicao:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                tse_temp_reposicao.append(sap_tse)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                tse_temp_reposicao.append(sap_tse)
                continue
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao

    @staticmethod
    def trocar_cv_kit():
        '''Módulo Pai Troca de Cavalete Kit'''
        etapa_reposicao = None
        identificador = Cavalete.TIPO
        tse_temp_reposicao = None
        tse_proibida = "TROCA CAVALETE (KIT)"
        # print("Iniciando processo Pai de Troca Cavalete KIT - TSE 149000")
        # servico_temp = session.findById(
        #     "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
        #     + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        # n_tse = 0
        # tse_temp_reposicao = []
        # num_tse_linhas = servico_temp.RowCount
        # print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        # for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
        #     sap_tse = servico_temp.GetCellValue(n_tse, "TSE")

        #     if sap_tse in tb_tse_reposicao:
        #         servico_temp.modifyCell(n_tse, "PAGAR", "n")
        #         # Pertence ao serviço principal
        #         servico_temp.modifyCell(n_tse, "CODIGO", "3")
        #         tse_temp_reposicao.append(sap_tse)
        #         continue

        #     elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
        #         servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
        #         # Pertence ao serviço principal
        #         servico_temp.modifyCell(n_tse, "CODIGO", "3")
        #         # Coloca a tse existente na lista temporária
        #         servico_temp.append(sap_tse)
        #         continue
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao

    @staticmethod
    def regularizar_cv():
        '''Módulo Pai Regularizar Cavalete.'''
        etapa_reposicao = None
        identificador = Cavalete.TIPO
        tse_temp_reposicao = None
        tse_proibida = "REGULARIZADO CAVALETE"
        # print("Iniciando processo Pai de Regularizar Cavalete - TSE 142000")
        # servico_temp = session.findById(
        #     "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
        #     + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        # n_tse = 0
        # tse_temp_reposicao = []
        # num_tse_linhas = servico_temp.RowCount
        # print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
        # for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
        #     sap_tse = servico_temp.GetCellValue(n_tse, "TSE")

        #     if sap_tse in tb_tse_reposicao:
        #         servico_temp.modifyCell(n_tse, "PAGAR", "n")
        #         # Pertence ao serviço principal
        #         servico_temp.modifyCell(n_tse, "CODIGO", "3")
        #         tse_temp_reposicao.append(sap_tse)
        #         continue

        #     elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
        #         servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
        #         # Pertence ao serviço principal
        #         servico_temp.modifyCell(n_tse, "CODIGO", "3")
        #         # Coloca a tse existente na lista temporária
        #         servico_temp.append(sap_tse)
        #         continue
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao

    @staticmethod
    def troca_cv_por_uma():
        '''Módulo Pai Troca Cavalete por UMA.'''
        identificador = Cavalete.TIPO
        etapa_reposicao = None
        tse_proibida = None
        print("Iniciando processo Pai de Troca de Cavalete por UMA - TSE 148000")
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
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                tse_temp_reposicao.append(sap_tse)
                etapa_reposicao = etapa
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                servico_temp.append(sap_tse)
                continue
        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao
