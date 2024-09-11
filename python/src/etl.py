"""Módulo para extração de ordens do SQL para a lista xlsx da Val."""

import typing
from pathlib import Path

import numpy as np
import pandas as pd

from python.src.sql_view import Sql

if typing.TYPE_CHECKING:
    from pandas import DataFrame


def pendentes_excel() -> np.ndarray:
    """Load de ordens em uma planilha expecífica."""
    path = Path.cwd() / "data"
    caminho = input("Digite o nome do arquivo.\n A planilha deve conter uma coluna Ordem\n")
    file_path = path / caminho
    planilha: DataFrame = pd.read_excel(f"{file_path}.XLSX", dtype=str)
    return planilha.to_numpy()


def pendentes_csv() -> list:
    """Load de ordens em uma planilha expecífica."""
    caminho = input("Digite o caminho do arquivo csv.\n O arquivo deve conter uma coluna Ordem\n")
    planilha = pd.read_csv(str(caminho))
    planilha = planilha.reset_index()
    return planilha["ORDEM"].to_list()


def extract_from_sql(contrato: str) -> np.ndarray:
    """Extração de ordens do contrato NOVASP do banco SQL Penha."""
    sql = Sql("", "")
    carteira = [
        "134000",
        "135000",
        "148000",
        "153000",
        "153500",
        "201000",
        "202000",
        "203000",
        "203500",
        "204000",
        "205000",
        "206000",
        "207000",
        "215000",
        # '253000'
        "254000",
        # '255000'
        # '262000',
        # '265000',
        # '266000':
        # '267000':
        # '268000':
        # '269000':
        # '280000',
        "284500",
        # '286000',
        # '304000'
        "322000",
        "405000",
        "406000",
        "407000",
        "408000",
        "414000",
        "450500",
        "453000",
        "455500",
        "463000",
        "465000",
        "467500",
        "475500",
        "502000",
        "505000",
        "506000",
        "508000",
        # '537000',
        # '537100',
        # '538000',
        # '561000'
        # '565000'
        "569000",
        # '581000'
        # '585000'
        # Serviços REM BASE
        "130000",
        "138000",
        "140000",
        "140100",
        "283000",
        "283500",
        "284000",
        "287000",
        "288000",
        "321000",
        "321500",
        # '325000',
        "328000",
        # '330000',
        "332000",
        "415000",
        "416000",
        "416500",
        # '540000'
        "531000",
        "531100",
        "560000",
        "567000",
        # '569000'
        "580000",
        "539000",
        "540000",
        "591000",
    ]
    return sql.carteira_tse(contrato, carteira)
