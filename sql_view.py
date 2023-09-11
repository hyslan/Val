'''Módulo para visualização da view de Valoração'''
import pandas as pd
from urllib.parse import quote_plus
from sqlalchemy import create_engine, text

server_name = '10.66.42.188'
database_name = 'BD_MLG'
connection_string = f'DRIVER={{SQL Server Native Client 11.0}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;'
encoded_connection_string = quote_plus(connection_string)
connection_url = f"mssql+pyodbc:///?odbc_connect={encoded_connection_string}"


class Tabela:
    '''Tabela de valoração'''

    def __init__(self, ordem, cod_tse) -> None:
        self.ordem = ordem
        self.cod_tse = cod_tse

    def tse_escolhida(self):
        '''Dados da tabela do SQL'''
        engine = create_engine(connection_url)
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
            sql_command = f"SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Engetami_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP IN ({tse}) AND [Feito?] IS NUll \
                AND MESREF >= {data_inicio} AND MESREF <= {data_fim}"
        else:
            sql_command = f"SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Engetami_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP IN ({tse}) AND [Feito?] IS NUll"

        df = pd.read_sql(sql_command, cnn)
        df = df.reset_index()
        df_list = df['Ordem'].tolist()
        return df_list

    def tse_expecifica(self):
        '''Dados da tabela do SQL'''
        engine = create_engine(connection_url)
        cnn = engine.connect()
        resposta = input(
            "- Val: Deseja escolher um período? \n")
        if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
            data_inicio = input(
                "- Val: Digite o Ano/Mês de ínicio, por favor.\n")
            data_fim = input(
                "- Val: Digite o Ano/Mês final, por favor.\n")
            sql_command = ("SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Engetami_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP = '{}' AND [Feito?] IS NUll \
            AND MESREF >= '{}' AND MESREF <= '{}'").format(self.cod_tse, data_inicio, data_fim)
        else:
            sql_command = ("SELECT Ordem FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Engetami_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP = '{}' AND [Feito?] IS NUll").format(self.cod_tse)
        df = pd.read_sql(sql_command, cnn)
        df = df.reset_index()
        df_list = df['Ordem'].tolist()
        return df_list

    def valorada(self, obs):
        '''Update de row valorada'''
        engine = create_engine(connection_url)
        cnn = engine.connect()
        quem = "Val"
        sql_command = ("INSERT INTO [LESTE_AD\hcruz_novasp].[tbHyslancruz_ENGETAMI_Valoradas]" +
                       "(Ordem, [VALORADO?], [POR QUEM?])" +
                       "VALUES ('{}', '{}', '{}')").format(str(self.ordem), obs, quem)
        cnn.execute(text(sql_command))
        cnn.commit()
        cnn.close()
