'''Módulo Família Poço Unitário.'''
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


class Poco:
    '''Classe Pai Hidrômetros.'''
    @staticmethod
    def det_descoberto_nivelado_reg_cx_parada():
        '''Módulo de Nivelar, Descobrir, Trocar Caixa de Parada'''
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
        num_tse_linhas = servico_temp.RowCount
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico_temp.GetCellValue(n_tse, "TSE")
            if sap_tse in tb_tse_PertenceAoServicoPrincipal:
                servico_temp.modifyCell(n_tse, "PAGAR", "n")
                # Pertence ao serviço principal
                servico_temp.modifyCell(n_tse, "CODIGO", "3")
                # Coloca a tse existente na lista temporária
                servico_temp.append(sap_tse)
                continue

        return reposicao_tse_temp, tse_proibida, identificador, etapa_reposicao
