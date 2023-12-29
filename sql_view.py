'''Módulo para visualização da view de Valoração'''
import pandas as pd
from urllib.parse import quote_plus
from sqlalchemy import create_engine, text


class Tabela:
    '''Tabela de valoração'''

    def __init__(self, ordem, cod_tse) -> None:
        self.ordem = ordem
        self.cod_tse = cod_tse
        server_name = '10.66.42.188'
        database_name = 'BD_MLG'
        connection_string = f'DRIVER={{SQL Server Native Client 11.0}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;'
        encoded_connection_string = quote_plus(connection_string)
        self.connection_url = f"mssql+pyodbc:///?odbc_connect={encoded_connection_string}"

    def tse_escolhida(self, contrato):
        '''Dados da tabela do SQL'''
        engine = create_engine(self.connection_url)
        cnn = engine.connect()
        tse = ','.join([f"'{tse}'" for tse in self.cod_tse])
        # Função desfazer valoração
        resposta = input(
            "- Val: Deseja escolher um período? \n")
        if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
            data_inicio = input(
                "- Val: Digite o Ano/Mês de ínicio, por favor.\n")
            data_fim = input(
                "- Val: Digite o Ano/Mês final, por favor.\n")
            sql_command = f"SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP IN ({tse}) AND [Feito?] IS NUll AND Contrato = '{contrato}' \
                AND MESREF >= {data_inicio} AND MESREF <= {data_fim}"
        else:
            sql_command = f"SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP IN ({tse}) AND [Feito?] IS NUll AND Contrato = '{contrato}'"

        df = pd.read_sql(sql_command, cnn)
        df = df.reset_index()
        df_list = df['Ordem'].tolist()
        cnn.close()
        return df_list

    def tse_expecifica(self, contrato):
        '''Dados da tabela do SQL'''
        engine = create_engine(self.connection_url)
        cnn = engine.connect()
        resposta = input(
            "- Val: Deseja escolher um período? \n")
        if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
            data_inicio = input(
                "- Val: Digite o Ano/Mês de ínicio, por favor.\n")
            data_fim = input(
                "- Val: Digite o Ano/Mês final, por favor.\n")
            sql_command = ("SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP = '{}' AND [Feito?] IS NUll AND Contrato = '{}' \
            AND MESREF >= '{}' AND MESREF <= '{}'").format(self.cod_tse, contrato, data_inicio, data_fim)
        else:
            sql_command = ("SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP = '{}' AND [Feito?] IS NUll AND Contrato = '{}'").format(self.cod_tse, contrato)

        df = pd.read_sql(sql_command, cnn)
        df = df.reset_index()
        df_list = df['Ordem'].tolist()
        cnn.close()
        return df_list

    def valorada(self, obs):
        '''Update de row valorada'''
        try:
            engine = create_engine(self.connection_url)
            cnn = engine.connect()
        except Exception as errosql:
            print(f"Erro SQL: {errosql}")
            engine = create_engine(self.connection_url)
            cnn = engine.connect()

        quem = "Val"
        sql_command = ("INSERT INTO [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Valoradas]" +
                       "(Ordem, [VALORADO?], [POR QUEM?])" +
                       "VALUES ('{}', '{}', '{}')").format(str(self.ordem), obs, quem)
        cnn.execute(text(sql_command))
        cnn.commit()
        cnn.close()

    def ordem_especifica(self, contrato):
        '''Teste de Ordem única.'''
        engine = create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = ("SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE ORDEM = '{}' AND Contrato = '{}'").format(str(self.ordem), contrato)
        df = pd.read_sql(sql_command, cnn)
        df = df.reset_index()
        df_list = df['Ordem'].tolist()
        cnn.close()
        return df_list

    def familia(self, familia, contrato):
        '''Escolher família.'''
        engine = create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = ("SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE FAMILIA = '{}' AND Contrato = '{}'").format(str(familia), contrato)
        df = pd.read_sql(sql_command, cnn)
        df = df.reset_index()
        df_list = df['Ordem'].tolist()
        cnn.close()
        return df_list

    def show_family(self):
        '''Print the family list.'''
        engine = create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = "SELECT [FAMILIA] FROM [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Parametros] \
            WHERE FAMILIA IS NOT NULL GROUP BY FAMILIA ORDER BY FAMILIA ASC "
        df = pd.read_sql(sql_command, cnn)
        # df = df.reset_index()
        df = df['FAMILIA'].to_string(index=False)
        cnn.close()
        return df

    def show_tses(self):
        '''Print the TSEs list.'''
        engine = create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = "SELECT COD_TSE, DESCRICAO FROM [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Parametros] \
            WHERE TP_PAGTO <> 'CANCELADO' AND COD_TSE IS NOT NULL ORDER BY COD_TSE ASC "
        df = pd.read_sql(sql_command, cnn)
        # df = df.reset_index()
        df = df.to_string(index=False)
        cnn.close()
        return df
