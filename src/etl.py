"""Módulo para extração de ordens do SQL para a lista xlsx da Val."""
from urllib.parse import quote_plus
import pandas as pd
from sqlalchemy import create_engine
from src.sql_view import Tabela

def pendentes_excel():
    """Load de ordens em uma planilha expecífica"""
    caminho = input(
        "Digite o caminho da planilha.\n A planilha deve conter uma coluna Ordem\n")
    planilha = pd.read_excel(str(caminho))
    lista = planilha.to_numpy()
    return lista


def pendentes_csv():
    '''Load de ordens em uma planilha expecífica'''
    caminho = input(
        "Digite o caminho do arquivo csv.\n O arquivo deve conter uma coluna Ordem\n")
    planilha = pd.read_csv(str(caminho))
    planilha = planilha.reset_index()
    lista = planilha['ORDEM'].to_list()
    return lista


def extract_from_sql(contrato):
    """Extração de ordens do contrato NOVASP do banco SQL Penha."""
    sql = Tabela("", "")
    carteira = [
        '134000',
        '135000',
        '148000',
        '153000',
        '153500',
        '201000',
        '202000',
        '203000',
        '203500',
        '204000',
        '205000',
        '206000',
        '207000',
        '215000',
        # '253000'
        '254000',
        # '255000'
        # '262000',
        # '265000',
        # '266000':
        # '267000':
        # '268000':
        # '269000':
        # '280000',
        '284500',
        # '286000',
        # '304000'
        '322000',
        '405000',
        '406000',
        '407000',
        '408000',
        '414000',
        '450500',
        '453000',
        '455500',
        '463000',
        '465000',
        '467500',
        '475500',
        '502000',
        '505000',
        '506000',
        '508000',
        # '537000',
        # '537100',
        # '538000',
        # '561000'
        # '565000'
        '569000',
        # '581000'
        # '585000'

        # Serviços REM BASE
        '130000',
        '138000',
        '140000',
        '140100',
        '283000',
        '283500',
        '284000',
        '287000',
        '288000',
        '321000',
        '321500',
        # '325000',
        '328000',
        # '330000',
        '332000',
        '415000',
        '416000',
        '416500',
        # '540000'
        '531000',
        '531100',
        '560000',
        '567000',
        # '569000'
        '580000',
        '539000',
        '540000',
        '591000',
    ]
    lista = sql.carteira_tse(contrato, carteira)
    return lista
