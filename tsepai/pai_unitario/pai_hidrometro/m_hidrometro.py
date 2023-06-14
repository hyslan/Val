# m_hidrometro.py
'''Módulo Família Hidrômetro Unitário.'''
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


class Hidrometro:
    '''Pai Hidrômetros.'''
    @staticmethod
    def UnitarioHidrometro(n_etapa):
        '''Módulo de hidrômetros, setar os parâmetros.'''
        print("Iniciando processo Pai de Hidromêtro Unitário - "
              + "TSE: 201000, 203000, 203500, 204000, 205000, 206000, 207000, 215000")
        servico_temp = session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
        n_tse = 0
        num_tse_linhas = servico_temp.RowCount
        n_etapa_hidro = servico_temp.GetCellValue(n_etapa, "ETAPA")
        for n_tse, SAP_tse in enumerate(range(0, num_tse_linhas)):
            SAP_tse = servico_temp.GetCellValue(n_tse, "TSE")

            if SAP_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                servico_temp.append(SAP_tse)
                continue
