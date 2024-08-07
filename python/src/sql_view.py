"""Módulo para visualização da view de Valoração"""
import os
from typing import Union
import datetime as dt
import numpy as np
import pandas as pd
import sqlalchemy as sa
from rich.console import Console
from dotenv import load_dotenv

console = Console()


class Sql:
    """Tabela de valoração"""

    def __init__(self, ordem: str, cod_tse: str) -> None:
        load_dotenv()
        self._ordem = ordem
        self.cod_tse = cod_tse
        connection_url = sa.URL.create(
            "mssql+pyodbc",
            username=os.environ['BD_USR'],
            password=os.environ['BD_PWD'],
            host=os.environ['BD_HOST'],
            database=os.environ['BD_NAME'],
            query={"driver": "ODBC Driver 18 for SQL Server",
                   "TrustServerCertificate": "yes",
                   },
        )
        self.connection_url = connection_url
        engine = sa.create_engine(self.connection_url)
        self.cnn = engine.connect()

    @property
    def ordem(self):
        return self._ordem

    @ordem.setter
    def ordem(self, cod):
        if isinstance(cod, str):
            self._ordem = cod
        else:
            raise ValueError("Wrong type, need to be string.")

    def __check_employee(self, matricula: str) -> str:
        """Check employee in the database:
        [LESTE_AD\\APP_Source].tb_Dim_Funcionarios
        """
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        query = ("SELECT NomeFuncionario FROM [LESTE_AD\\APP_Source].[tb_Dim_Funcionarios] "
                 f"WHERE Matricula = '{matricula}';")
        df = pd.read_sql(query, cnn)
        who: str = df.to_string(index=False)
        cnn.close()
        return who

    def carteira_tse(self, contrato, carteira):
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        carteira_str = ','.join([f"'{tse}'" for tse in carteira])
        # Queries para SQL.
        # pylint disable=W1401
        query = (f"SELECT [Ordem], COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                 f"WHERE [TSE_OPERACAO_ZSCP] IN ({carteira_str}) "
                 f"AND Contrato = '{contrato}';")
        df = pd.read_sql(query, cnn)
        df_array = df.to_numpy()
        cnn.close()
        print("\nExtração de ordens feita com sucesso!")
        return df_array

    def tse_escolhida(self, contrato):
        """Dados da tabela do SQL"""
        engine = sa.create_engine(self.connection_url)
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
            sql_command = f"SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP IN ({tse}) AND Contrato = '{contrato}' \
                AND MESREF >= {data_inicio} AND MESREF <= {data_fim}"
        else:
            sql_command = f"SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] \
            WHERE TSE_OPERACAO_ZSCP IN ({tse}) AND Contrato = '{contrato}'"

        df = pd.read_sql(sql_command, cnn)
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def tse_expecifica(self, contrato):
        """Dados da tabela do SQL"""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        resposta = input(
            "- Val: Deseja escolher um período? \n")
        if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
            data_inicio = input(
                "- Val: Digite o Ano/Mês de ínicio, por favor.\n")
            data_fim = input(
                "- Val: Digite o Ano/Mês final, por favor.\n")
            sql_command = ("SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                           f"WHERE TSE_OPERACAO_ZSCP = '{self.cod_tse}' AND "
                           f"Contrato = '{contrato}' "
                           f"AND MESREF >= '{data_inicio}' AND MESREF <= '{data_fim}'")
        else:
            sql_command = ("SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                           f"WHERE TSE_OPERACAO_ZSCP = '{self.cod_tse}' AND "
                           f"Contrato = '{contrato}'")

        df = pd.read_sql(sql_command, cnn)
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def clean_duplicates(self):
        """Delete duplicates rows
        Keeping only the most recent one"""
        try:
            engine = sa.create_engine(self.connection_url)
            cnn = engine.connect()
        except Exception as errosql:
            print(f"Erro SQL: {errosql}")
            engine = sa.create_engine(self.connection_url)
            cnn = engine.connect()

        sql_command = ("WITH CTE AS ("
                       "SELECT *,"
                       " ROW_NUMBER() OVER(PARTITION BY Ordem ORDER BY DataRegistro DESC) AS RowNumber"
                       " FROM [LESTE_AD\\hcruz_novasp].tbHyslancruz_Valoradas"
                       ")"
                       " DELETE FROM CTE WHERE RowNumber > 1;"
                       )
        cnn.execute(sa.text(sql_command))
        cnn.commit()
        cnn.close()

    def retrabalho_search(self, month_start: Union[str | dt.date], month_end: Union[str | dt.date]) -> np.ndarray:
        """Query for Retrabalho confirmado orders
        Args:
            month_start Union[str | Date]: Start month of the query
            month_end Union[str | Date]: End month of the query
            Not Returning SABESP -> Cod: '9999999999' orders
        Returns:
            df_array (np.ndarray): Array with the query results
            """
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()

        sql_command = ("SELECT NumeroOS, ATC, CodigoContrato FROM [LESTE_AD\\CargaDeDados].[tb_Fato_Bexec] "
                       f"WHERE DataFimExecucao >= '{
                           month_start}' AND DataFimExecucao <= '{month_end}' "
                       "AND Resultado = 'RETRABALHO CONFIRMADO' AND CodigoContrato <> '9999999999'")
        df = pd.read_sql(sql_command, cnn)
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def valorada(self, valorado: str, contrato: str,
                 municipio: str, status: str, obs: str,
                 data_valoracao: Union[None | dt.date | str],
                 matricula: str, valor_medido: float, tempo_gasto: float) -> None:
        """Update  row valorada to 
        [LESTE_AD\\hcruz_novasp].tbHyslancruz_Valoradas
        """
        try:
            engine = sa.create_engine(self.connection_url)
            cnn = engine.connect()
        except Exception as errosql:
            print(f"Erro SQL: {errosql}")
            engine = sa.create_engine(self.connection_url)
            cnn = engine.connect()

        if matricula == '117615':
            quem = "Val"
        else:
            # ? quem = self.__check_employee(matricula)
            quem = 'teste'

        if data_valoracao is None:
            data_valoracao = dt.datetime.now().date()
        else:
            data_valoracao = data_valoracao.replace('.', '-')
            data_valoracao = dt.datetime.strptime(
                data_valoracao, '%d-%m-%Y').date()

        data_valoracao = data_valoracao.strftime('%m/%d/%Y')

        sql_command = ("INSERT INTO [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Valoradas] " +
                       "(Ordem, [VALORADO?], [POR QUEM?], Contrato, TSE, Municipio, Status, " +
                       "OBS, TempoGasto, DataValoracao, Matricula, VALOR_MEDIDO)" +
                       f"VALUES ('{self.ordem}', '{valorado}', '{quem}', '{contrato}', " +
                       f"'{self.cod_tse}', '{municipio}', '{status}', '{obs}', " +
                       f"'{tempo_gasto}', '{data_valoracao}', '{matricula}', '{valor_medido}')")
        cnn.execute(sa.text(sql_command))
        cnn.commit()
        cnn.close()

    def ordem_especifica(self, contrato):
        """Teste de Ordem única."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = ("SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                       f"WHERE ORDEM = '{str(self.ordem)}' AND Contrato = '{contrato}'")
        df = pd.read_sql(sql_command, cnn)
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def familia(self, family: Union[list[str] | None], contrato: str) -> np.ndarray:
        """Escolher família."""
        if family is not None:
            family_str = ','.join([f"'{f}'" for f in family])
        else:
            family_str = ("'CAVALETE', 'HIDROMETRO', 'POCO', 'RAMAL AGUA', 'RELIGACAO', 'SUPRESSAO' "
                          # TODO: Resolve each TSE of these families below:
                          # "'REDE AGUA'"  # , 'REDE ESGOTO', 'RAMAL ESGOTO'," <- Sem tubo dn 100
                          )

        console.print("\n [b]Família escolhida: ", family_str)
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        # TSEs leave out the plumbing services by Iara.
        chief_iara_orders = "'534200', '534300', '537000', '537100', '538000'," if contrato == "4600042975" else "'',"
        sql_command = fr"""
            SELECT Ordem, COD_MUNICIPIO
            FROM [LESTE_AD\hcruz_novasp].[v_Hyslan_Valoracao]
            WHERE FAMILIA IN ({family_str})
            AND Contrato = '{contrato}'
            AND TSE_OPERACAO_ZSCP NOT IN (
                '731000', '733000', '743000', '745000', '785000', '785500',
                '755000', '714000', '782500', '282000', '300000', '308000', '310000', '311000', '313000',
                '315000', '532000', '564000', '588000', '590000', '709000', '700000', '593000', '253000',
                '250000', '209000', '605000', '605000', '263000', '255000', '254000', '282000', '265000',
                '260000', '265000', '263000', '262000', '284500', '286000', '282500',
                {chief_iara_orders}
                '136000', '159000', '155000');
        """

        console.print(f"\n[bold yellow]{sql_command}")
        df = pd.read_sql(sql_command, cnn)
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def show_family(self):
        """Print the family list."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = "SELECT [FAMILIA] FROM [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Parametros] \
            WHERE FAMILIA IS NOT NULL GROUP BY FAMILIA ORDER BY FAMILIA"
        df = pd.read_sql(sql_command, cnn)
        df = df['FAMILIA'].to_string(index=False)
        cnn.close()
        return df

    def show_tses(self):
        """Print the TSEs list."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = ("SELECT COD_TSE, DESCRICAO FROM "
                       "[LESTE_AD\\hcruz_novasp].[tbHyslancruz_Parametros] "
                       "WHERE TP_PAGTO <> 'CANCELADO' AND "
                       "COD_TSE IS NOT NULL ORDER BY COD_TSE")
        df = pd.read_sql(sql_command, cnn)
        df = df.to_string(index=False)
        cnn.close()
        return df

    def get_new_hidro(self):
        """Get new hidro from SQL."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = ("SELECT HidrometroInstalado FROM "
                       "[LESTE_AD\\CargaDeDados].tb_Fato_BexecHidros "
                       f"WHERE NumeroOS = '{self.ordem}'")
        df = pd.read_sql(sql_command, cnn)
        df = df.to_string(index=False)
        cnn.close()
        return df
