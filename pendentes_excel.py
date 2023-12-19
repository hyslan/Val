'''Carregar planilha pendente'''
import pandas as pd


def pendentes():
    '''Load de ordens em uma planilha expecífica'''
    caminho = input(
        "Digite o caminho da planilha.\n A planilha deve conter uma coluna Ordem\n")
    planilha = pd.read_excel(caminho)
    planilha = planilha.reset_index()
    lista = planilha['Ordem'].to_list()
    return lista
