#Excel_tbs.py by Author: Hyslan Silva Cruz
from openpyxl import load_workbook # Carregar função load
#Área do Excel
def load_worksheets():
    lista = load_workbook('lista.xlsx') # Carregando arquivo para valorar
    materiais = load_workbook('MateriaisContratada.xlsx') # Tabela de materiais da Contratada
    plan_tse = load_workbook('TSE.xlsx') # Tabela de TSE
    plan_precos = load_workbook('ItensPreco.xlsx') # Tabela Itens de preço
    plan_Stordem = load_workbook('StatusOrdem.xlsx') # Tabela Status da Ordem
    planilha = lista.active # Planilha de Ordem
    contratada = materiais.active  # Materiais Contratada
    StSistema = plan_Stordem["StatusSistema"] # Status Sistema
    StUsuario = plan_Stordem["StatusUsuario"] # Status Usuario
    unitario = plan_tse["unitario"] 
    rb = plan_tse["rb"]
    naoexecutado = plan_tse["NEXEC"]
    invest = plan_tse["invest"]
    n3 = plan_tse["n3"]
    n10 = plan_tse["n10"]
    reposicao = plan_tse["reposicao"]
    retrabalho = plan_tse["retrabalho"]
    asfalto = plan_tse["asfalto"]
    coluna_contratada = 'A'
    coluna_tse = 'A'
    coluna_status = 'A'
    
    tb_contratada = [cell.value for cell in contratada[coluna_contratada]]
    
    plan_Stordem.active = StSistema
    tb_StSistema = [cell.value for cell in StSistema[coluna_status]]
    
    plan_Stordem.active = StUsuario
    tb_StUsuario = [cell.value for cell in StSistema[coluna_status]]

    plan_tse.active = unitario
    tb_tse_UN = [cell.value for cell in unitario[coluna_tse]]

    plan_tse.active = rb
    tb_tse_rb = [cell.value for cell in rb[coluna_tse]]

    plan_tse.active = naoexecutado
    tb_tse_nexec = [cell.value for cell in naoexecutado[coluna_tse]]

    plan_tse.active = invest
    tb_tse_invest = [cell.value for cell in invest[coluna_tse]]
    
    plan_tse.active = n3
    tb_tse_PertenceAoServicoPrincipal = [cell.value for cell in n3[coluna_tse]]
    
    plan_tse.active = n10
    tb_tse_ServicoNaoExistenoContrato = [cell.value for cell in n10[coluna_tse]]
    
    plan_tse.active = reposicao
    tb_tse_reposicao = [cell.value for cell in reposicao[coluna_tse]]
    
    plan_tse.active = retrabalho
    tb_tse_Retrabalho = [cell.value for cell in retrabalho[coluna_tse]]
    
    plan_tse.active = asfalto
    tb_tse_Asfalto = [cell.value for cell in asfalto[coluna_tse]]

    return (
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
        tb_tse_Asfalto
    )