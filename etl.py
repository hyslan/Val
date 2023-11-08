'''Módulo para extração de ordens do SQL para a lista xlsx da Val.'''
import pandas as pd
from urllib.parse import quote_plus
from sqlalchemy import create_engine, text


def extract_from_sql(contrato):
    '''Extração de ordens do contrato NOVASP do banco SQL Penha.'''
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
        '416000',
        # '539000'
        # '540000'
        '560000',
        '567000',
        # '569000'
        '580000',
        '539000',
        '540000',
        '591000',
    ]

    # Construção da cláusula IN como uma string separada por vírgulas
    carteira_str = ','.join([f"'{tse}'" for tse in carteira])
    server_name = '10.66.42.188'
    database_name = 'BD_MLG'
    connection_string = f'DRIVER={{SQL Server Native Client 11.0}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;'
    encoded_connection_string = quote_plus(connection_string)
    connection_url = f"mssql+pyodbc:///?odbc_connect={encoded_connection_string}"
    engine = create_engine(connection_url)
    cnxn = engine.connect()
    print("\nConexão com SQL bem sucedida.\n")

    # Queries para SQL.

    # pylint disable=W1401
    QUERY = f"SELECT [Ordem] FROM [BD_MLG].[LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE [TSE_OPERACAO_ZSCP] IN ({carteira_str}) AND [Feito?] IS NUll AND Contrato = '{contrato}';"
    df = pd.read_sql(QUERY, cnxn)
    pendentes = pd.DataFrame(df)
    pendentes_list = pendentes['Ordem'].tolist()
    print("\nExtração de ordens feita com sucesso!")
    return pendentes_list
