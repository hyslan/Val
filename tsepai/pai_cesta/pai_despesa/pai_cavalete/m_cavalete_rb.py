# cavalete.py
'''Módulo família Cavalete - remuneração base '''
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


class Cavalete:
    '''Familia Cavalete'''
    MODALIDADE = "ligacao_agua"

    @staticmethod
    def reparo_cv():
        '''Reparo de Cavalete'''
        etapa_reposicao = None
        identificador = Cavalete.MODALIDADE
        print("Iniciando processo Pai de Reparo de Cavalete - TSE 130000")
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
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                tse_temp_reposicao.append(sap_tse)
                continue
        return tse_temp_reposicao, identificador, etapa_reposicao

    @staticmethod
    def reparo_de_registro_de_cv():
        '''Reparo de Registro de Cavalete.'''
        etapa_reposicao = None
        identificador = Cavalete.MODALIDADE
        print("Iniciando processo Pai de Reparo de Registro de Cavalete - TSE 140000")
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
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                tse_temp_reposicao.append(sap_tse)
                continue
        return tse_temp_reposicao, identificador, etapa_reposicao

    @staticmethod
    def troca_de_registro_de_cv():
        '''Troca de registro de Cavalete.'''
        etapa_reposicao = None
        identificador = Cavalete.MODALIDADE
        print("Iniciando processo Pai de Troca de Registro de Cavalete - TSE 140100")
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
                servico_temp.modifyCell(n_tse, "CODIGO", "5")  # Despesa
                tse_temp_reposicao.append(sap_tse)
                continue

            elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                tse_temp_reposicao.append(sap_tse)
                continue
        return tse_temp_reposicao, identificador, etapa_reposicao
