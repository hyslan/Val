"""Módulo de testes de SQL View.

Conexões e queries.
"""

import os

import pandas as pd
import sqlalchemy as sa
from dotenv import load_dotenv
from sqlalchemy.exc import SQLAlchemyError


def setup_connection() -> sa.Connection:
    """Set up SQL Connection."""
    load_dotenv()
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
    engine = sa.create_engine(connection_url)
    return engine.connect()


def test_connection() -> None:
    """Test SQL Connection."""
    conn = setup_connection()
    # Check if the connection is active
    if conn.connection.is_valid:
        assert conn.connection.is_valid
    else:
        raise SQLAlchemyError


def test_query_familia() -> None:
    """Test SQL Query."""
    family_str = "HIDROMETRO, RELIGACAO"
    contrato = "4600041302"
    # chief_iara_orders = "'534200', '534300', '537000', '537100', '538000'," if contrato == "4600042975" else ""
    conn = setup_connection()
    query = sa.text(r"""
                    SELECT Ordem, COD_MUNICIPIO
            FROM [LESTE_AD\hcruz_novasp].[v_Hyslan_Valoracao]
            WHERE FAMILIA IN (:family_str)
            AND Contrato = :contrato
            AND [Feito?] NOT IN ('SIM', 'Num Pode', N'Sem posição de rede.', 'Definitiva')
            AND TSE_OPERACAO_ZSCP NOT IN (
                '731000', '733000', '743000', '745000', '785000', '785500',
                '755000', '714000', '782500', '282000', '300000', '308000', '310000', '311000', '313000',
                '315000', '532000', '564000', '588000', '590000', '709000', '700000', '593000', '253000',
                '250000', '209000', '605000', '605000', '263000', '255000', '254000', '282000', '265000',
                '260000', '265000', '263000', '262000', '284500', '286000', '282500',
                '136000', '159000', '155000');""")

    df = pd.read_sql(
        query,
        conn,
        params={"contrato": contrato, "family_str": family_str},
    )
    df_array = df.to_numpy()
    assert df_array is not None
