# ServicosExecutados.py
'''Módulo de TSE'''
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from tsepai import pai_dicionario

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


def verifica_tse(servico):
    '''Agrupador de serviço e indexador de classes.'''
    print("Iniciando processo de verificação de TSE")
    servico = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                               + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
    n_tse = 0
    tse_temp = []  # Lista temporária para armazenar as tse
    num_tse_linhas = servico.RowCount
    print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
    for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
        sap_tse = servico.GetCellValue(n_tse, "TSE")
        if sap_tse in tb_tse_UN:  # Verifica se está no Conjunto Unitários
            servico.modifyCell(n_tse, "PAGAR", "s")  # Marca pagar na TSE
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            reposicao = pai_dicionario.pai_servico_unitario(sap_tse)
            continue
        elif sap_tse in tb_tse_rb:  # Caso Contrário, é RB - Despesa
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5")  # Despesa
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            pai_dicionario.pai_servico_cesta(sap_tse)
            continue
          # tirar depois
        elif sap_tse in tb_tse_invest:  # Caso Contrário, é RB - Investimento
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "6")  # Investimento
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            continue
            # tirar depois
        elif sap_tse in tb_tse_nexec:
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "11")  # Não Executado
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            continue
        elif sap_tse in tb_tse_PertenceAoServicoPrincipal:
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Pertence ao Serviço Principal
            servico.modifyCell(n_tse, "CODIGO", "3")
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            continue
        elif sap_tse in tb_tse_Retrabalho:
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "7")  # Retrabalho
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            continue
        else:
            print("TSE não encontrado na planilha do Excel.")
    # Fim da condicional.
    servico.pressEnter()
    return tse_temp, reposicao, num_tse_linhas
