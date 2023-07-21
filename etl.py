'''Módulo para extração de ordens do SQL para a lista xlsx da Val.'''
import pandas as pd
import pyodbc


def extract_from_sql():
    '''Extração de ordens do contrato NOVASP do banco SQL Penha.'''
    carteira = [
        '138000',
        '142000',
        '148000',
        '149000',
        '153000',
        '201000',
        '202000',
        '203000',
        '203500',
        '204000',
        '205000',
        '206000',
        '207000',
        '211000',
        '215000',
        # '253000'
        # '254000'
        # '255000'
        # '262000'
        # '265000'
        # '266000':
        # '267000':
        # '268000':
        # '269000':
        # '284500'
        # '286000'
        # '304000'
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
        # '502000'
        # '505000'
        # '506000'
        # '508000'
        # '537000'
        # '537100'
        # '538000'
        # '561000'
        # '565000'
        # '569000'
        # '581000'
        # '585000'
        # '713000'
        # '713500'

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
        # '325000'
        # '328000'
        # '330000'
        '332000',
        '416000',
        # '539000'
        # '540000'
        '560000',
        '567000',
        # '569000'
        '580000',
        # '539000'
        # '540000'
        '591000',
    ]

    # Construção da cláusula IN como uma string separada por vírgulas
    carteira_str = ','.join([f"'{tse}'" for tse in carteira])
    SERVER = '10.66.9.46'
    DATABASE = 'BD_MLG'
    cnxn = pyodbc.connect('DRIVER={SQL SERVER};SERVER='+SERVER +
                          ';DATABASE='+DATABASE+";Integrated Security=SSPI;")
    print("\nConexão com SQL bem sucedida.\n")
    cursor = cnxn.cursor()
    # Queries para SQL.

    # pylint disable=W1401
    QUERY = f"SELECT [Ordem] FROM [BD_MLG].[LESTE_AD\hcruz_novasp].[v_Hyslan_Engetami_Valoracao] \
            WHERE [TSEOperacaoZSCP] IN ({carteira_str});"
    df = pd.read_sql(QUERY, cnxn)
    pendentes = pd.DataFrame(df)
    ORDENS = 'lista.xlsx'
    pendentes.to_excel(ORDENS, index=False, header=False)
    print("\nExtração de ordens feita com sucesso!")
