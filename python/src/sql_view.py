"""Módulo para visualização da view de Valoração."""

import datetime as dt
import logging
import os
from typing import Any

import numpy.typing as npt
import pandas as pd
import pytz
import sqlalchemy as sa
from dotenv import load_dotenv
from rich.console import Console
from sqlalchemy.exc import SQLAlchemyError

console = Console()
logger = logging.getLogger(__name__)


class Sql:
    """Tabela de valoração."""

    def __init__(self, ordem: str, cod_tse: str | list[str]) -> None:
        """Inicializa a conexão com o banco de dados.

        Args:
        ----
            ordem (str): Número da Ordem
            cod_tse (str | list[str]): Código da TSE

        """
        load_dotenv()
        self._ordem = ordem
        self.cod_tse = cod_tse
        connection_url = sa.URL.create(
            "mssql+pyodbc",
            username=os.environ["BD_USR"],
            password=os.environ["BD_PWD"],
            host=os.environ["BD_HOST"],
            database=os.environ["BD_NAME"],
            query={
                "driver": "ODBC Driver 18 for SQL Server",
                "TrustServerCertificate": "yes",
            },
        )
        self.connection_url = connection_url
        engine = sa.create_engine(self.connection_url)
        self.cnn = engine.connect()

    @property
    def ordem(self) -> str:
        """Get the ordem value.

        Returns
        -------
            str: Número da Ordem

        """
        return self._ordem

    @ordem.setter
    def ordem(self, cod: str) -> None:
        if isinstance(cod, str):
            self._ordem = cod

    def __check_employee(self, matricula: str) -> str:
        r"""Check employee in the database.

        [LESTE_AD\\APP_Source].tb_Dim_Funcionarios.
        """
        engine = sa.create_engine(self.connection_url)
        with engine.connect() as cnn:
            query = sa.text(
                "SELECT NomeFuncionario FROM [LESTE_AD\\APP_Source].[tb_Dim_Funcionarios] WHERE Matricula = :matricula",
            )
            df = pd.read_sql(query, cnn, params={"matricula": matricula})
            who: str = df.to_string(index=False)

        return who

    def carteira_tse(self, contrato: str, carteira: list[str]) -> npt.NDArray[Any]:
        """Dados da tabela do SQL.

        Args:
        ----
            contrato (str): Número do Contrato
            carteira (list[str]): Lista do etl.py

        Returns:
        -------
            np.ndarray: Array com os resultados da query (ordem, código do município).

        """
        engine = sa.create_engine(self.connection_url)
        carteira_str = ",".join([f"'{tse}'" for tse in carteira])
        # Queries para SQL.
        with engine.connect() as cnn:
            query = sa.text(
                "SELECT [Ordem], COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                "WHERE [TSE_OPERACAO_ZSCP] IN :carteira "
                "AND Contrato = :contrato",
            )
            df = pd.read_sql(query, cnn, params={"carteira": carteira_str, "contrato": contrato})
            return df.to_numpy()

    def tse_escolhida(self, contrato: str) -> npt.NDArray[Any]:
        """Dados da tabela do SQL."""
        engine = sa.create_engine(self.connection_url)
        tse = ",".join([f"'{tse}'" for tse in self.cod_tse])
        # Função desfazer valoração
        resposta = input("- Val: Deseja escolher um período? \n")
        if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
            data_inicio = input("- Val: Digite o Ano/Mês de ínicio, por favor.\n")
            data_fim = input("- Val: Digite o Ano/Mês final, por favor.\n")
            sql_command = sa.text(
                "SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                "WHERE TSE_OPERACAO_ZSCP IN :tse AND Contrato = :contrato "
                "AND MESREF >= :datainicio AND MESREF <= :datafim",
            )
            with engine.connect() as cnn:
                df = pd.read_sql(
                    sql_command,
                    cnn,
                    params={"tse": tse, "contrato": contrato, "datainicio": data_inicio, "datafim": data_fim},
                )
                return df.to_numpy()
        else:
            sql_command = sa.text(
                "SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                "WHERE TSE_OPERACAO_ZSCP IN :tse AND Contrato = :contrato",
            )
            with engine.connect() as cnn:
                df = pd.read_sql(sql_command, cnn, params={"tse": tse, "contrato": contrato})
                return df.to_numpy()

    def tse_expecifica(self, contrato: str) -> npt.NDArray[Any]:
        """Dados da tabela do SQL."""
        engine = sa.create_engine(self.connection_url)
        tse = ",".join([f"'{tse}'" for tse in self.cod_tse])
        resposta = input("- Val: Deseja escolher um período? \n")
        if resposta in ("s", "S", "sim", "Sim", "SIM", "y", "Y", "yes"):
            data_inicio = input("- Val: Digite o Ano/Mês de ínicio, por favor.\n")
            data_fim = input("- Val: Digite o Ano/Mês final, por favor.\n")
            sql_command = sa.text(
                "SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                "WHERE TSE_OPERACAO_ZSCP = :codtse AND "
                "Contrato = :contrato "
                "AND MESREF >= :datainicio AND MESREF <= :datafim",
            )

            with engine.connect() as cnn:
                df = pd.read_sql(
                    sql_command,
                    cnn,
                    params={"codtse": tse, "contrato": contrato, "datainicio": data_inicio, "datafim": data_fim},
                )
                return df.to_numpy()

        else:
            sql_command = sa.text(
                "SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
                "WHERE TSE_OPERACAO_ZSCP = :codtse AND "
                "Contrato = :contrato",
            )

            with engine.connect() as cnn:
                df = pd.read_sql(sql_command, cnn, params={"codtse": tse, "contrato": contrato})
                return df.to_numpy()

    def clean_duplicates(self) -> None:
        """Delete duplicates rows.

        Keeping only the most recent one.
        """
        try:
            engine = sa.create_engine(self.connection_url)
            cnn = engine.connect()
        except SQLAlchemyError:
            logger.exception("Erro ao conectar com o banco de dados em Clean Duplicates")

        sql_command = (
            "WITH CTE AS ("
            "SELECT *,"
            " ROW_NUMBER() OVER(PARTITION BY Ordem ORDER BY DataRegistro DESC) AS RowNumber"
            " FROM [LESTE_AD\\hcruz_novasp].tbHyslancruz_Valoradas"
            ")"
            " DELETE FROM CTE WHERE RowNumber > 1;"
        )
        cnn.execute(sa.text(sql_command))
        cnn.commit()
        cnn.close()

    def retrabalho_search(
        self,
        month_start: str | dt.date,
        month_end: str | dt.date,
    ) -> npt.NDArray[Any]:
        """Query for Retrabalho confirmado orders.

        Not Returning SABESP -> Cod: '9999999999' orders

        Args:
        ----
            month_start (str | dt.date): Start month of the query
            month_end (str | dt.date): End month of the query

        Returns:
        -------
            df_array (np.ndarray): Array with the query results.

        """
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()

        sql_command = sa.text(
            "SELECT NumeroOS, ATC, CodigoContrato FROM [LESTE_AD\\CargaDeDados].[tb_Fato_Bexec] "
            "WHERE DataFimExecucao >= :monthstart AND DataFimExecucao <= :monthend "
            "AND Resultado = 'RETRABALHO CONFIRMADO' AND CodigoContrato <> '9999999999'",
        )
        df = pd.read_sql(sql_command, cnn, params={"monthstart": month_start, "monthend": month_end})
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def valorada(
        self,
        valorado: str,
        contrato: str,
        municipio: str,
        status: str,
        obs: str,
        data_valoracao: None | dt.date | str,
        matricula: str,
        valor_medido: float,
        tempo_gasto: float,
    ) -> None:
        r"""Update  row valorada to.

        [LESTE_AD\\hcruz_novasp].tbHyslancruz_Valoradas.
        """
        try:
            engine = sa.create_engine(self.connection_url)
            cnn = engine.connect()
        except SQLAlchemyError:
            logger.exception("Erro ao conectar com o banco de dados em Valorada")

        quem = "Val" if matricula == "117615" else self.__check_employee(matricula)

        if data_valoracao is None:
            data_valoracao = dt.datetime.now().date()
        elif isinstance(data_valoracao, str):
            data_valoracao = data_valoracao.replace(".", "-")
            data_valoracao = (
                dt.datetime.strptime(data_valoracao, "%d-%m-%Y").astimezone(pytz.timezone("America/Sao_Paulo")).date()
            )

        data_valoracao = data_valoracao.strftime("%m/%d/%Y")

        sql_command = (
            "INSERT INTO [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Valoradas] "
            "(Ordem, [VALORADO?], [POR QUEM?], Contrato, TSE, Municipio, Status, "
            "OBS, TempoGasto, DataValoracao, Matricula, VALOR_MEDIDO)"
            f"VALUES ('{self.ordem}', '{valorado}', '{quem}', '{contrato}', "
            f"'{self.cod_tse}', '{municipio}', '{status}', '{obs}', "
            f"'{tempo_gasto}', '{data_valoracao}', '{matricula}', '{valor_medido}')"
        )
        try:
            cnn.execute(sa.text(sql_command))
            cnn.commit()
        except SQLAlchemyError:
            logger.exception("Erro ao executar a query em Valorada.")
        cnn.close()

    def ordem_especifica(self, contrato: str) -> npt.NDArray[Any]:
        """Apenas uma ordem específica.

        Args:
        ----
            contrato (str): Número do Contrato

        Returns:
        -------
            np.ndarray: Array com os resultados da query (ordem, código do município).

        """
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = sa.text(
            "SELECT Ordem, COD_MUNICIPIO FROM [LESTE_AD\\hcruz_novasp].[v_Hyslan_Valoracao] "
            "WHERE ORDEM = :ordem AND Contrato = :contrato'",
        )
        df = pd.read_sql(sql_command, cnn, params={"ordem": self.ordem, "contrato": contrato})
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def familia(self, family: list[str] | None, contrato: str) -> npt.NDArray[Any]:
        """Escolher família."""
        if family is not None:
            family_str = ",".join([f"'{f}'" for f in family])
        else:
            family_str = (
                "'CAVALETE', 'HIDROMETRO', 'POCO', 'RAMAL AGUA', 'RELIGACAO', 'SUPRESSAO' "
                "'REDE AGUA', 'REDE ESGOTO', 'RAMAL ESGOTO',"
            )

        console.print("\n [b]Família escolhida: ", family_str)
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        # TSEs leave out the plumbing services by Iara.
        chief_iara_orders = "'534200', '534300', '537000', '537100', '538000'," if contrato == "4600042975" else "'',"
        sql_command = sa.text(r"""
            SELECT Ordem, COD_MUNICIPIO
            FROM [LESTE_AD\hcruz_novasp].[v_Hyslan_Valoracao]
            WHERE FAMILIA IN :family_str
            AND Contrato = '{contrato}'
            AND TSE_OPERACAO_ZSCP NOT IN (
                '731000', '733000', '743000', '745000', '785000', '785500',
                '755000', '714000', '782500', '282000', '300000', '308000', '310000', '311000', '313000',
                '315000', '532000', '564000', '588000', '590000', '709000', '700000', '593000', '253000',
                '250000', '209000', '605000', '605000', '263000', '255000', '254000', '282000', '265000',
                '260000', '265000', '263000', '262000', '284500', '286000', '282500',
                :chief_iara_orders
                '136000', '159000', '155000');
        """)

        console.print(f"\n[bold yellow]{sql_command}")
        df = pd.read_sql(
            sql_command,
            cnn,
            params={"family_str": family_str, "contrato": contrato, "chief_iara_orders": chief_iara_orders},
        )
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def desobstrucao(self) -> npt.NDArray[Any]:
        """Desobstrução NORTESUL."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = r"""
            SELECT Ordem, COD_MUNICIPIO, CONTRATO
            FROM [LESTE_AD\hcruz_novasp].[v_Hyslan_Valoracao]
            WHERE Contrato IN ('4600043760', '4600045267', '4600046036', '4600046036');
        """
        console.print(f"\n[bold yellow]{sql_command}")
        df = pd.read_sql(sql_command, cnn)
        df_array = df.to_numpy()
        cnn.close()
        return df_array

    def show_family(self) -> str:
        """Print the family list."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = "SELECT [FAMILIA] FROM [LESTE_AD\\hcruz_novasp].[tbHyslancruz_Parametros] \
            WHERE FAMILIA IS NOT NULL GROUP BY FAMILIA ORDER BY FAMILIA"
        df = pd.read_sql(sql_command, cnn)
        list_family = df["FAMILIA"].to_string(index=False)
        cnn.close()
        return list_family

    def show_tses(self) -> str:
        """Print the TSEs list."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = (
            "SELECT COD_TSE, DESCRICAO FROM "
            "[LESTE_AD\\hcruz_novasp].[tbHyslancruz_Parametros] "
            "WHERE TP_PAGTO <> 'CANCELADO' AND "
            "COD_TSE IS NOT NULL ORDER BY COD_TSE"
        )
        df = pd.read_sql(sql_command, cnn)
        list_tses = df.to_string(index=False)
        cnn.close()
        return list_tses

    def get_new_hidro(self) -> str:
        """Get new hidro from SQL."""
        engine = sa.create_engine(self.connection_url)
        cnn = engine.connect()
        sql_command = sa.text(
            "SELECT HidrometroInstalado FROM [LESTE_AD\\CargaDeDados].tb_Fato_BexecHidros WHERE NumeroOS = :ordem",
        )
        df = pd.read_sql(sql_command, cnn, params={"ordem": self.ordem})
        hidro = df.to_string(index=False)
        cnn.close()
        return hidro
