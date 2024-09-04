import pandas as pd
import pytest
from pandas import DataFrame

from src.wms.localiza_material import qtd_max


@pytest.fixture
def setup_data() -> tuple[DataFrame, DataFrame]:
    """Fixture para setup de dados.

    Returns
    -------
        tuple[DataFrame, DataFrame]: Dataframes

    """
    estoque = pd.DataFrame(
        {
            "Material": ["A", "B", "C"],
            "Quantidade": [100, 200, 300],
        },
    )
    df_materiais = pd.DataFrame(
        {
            "Material": ["A", "A", "B", "C", "C"],
            "Quantidade": [50, 150, 250, 350, 450],
        },
    )
    return estoque, df_materiais


def test_material_no_estoque(setup_data: tuple[DataFrame, DataFrame]) -> None:
    """Testa se a quantidade máxima é retornada corretamente.

    Args:
    ----
        setup_data (tuple[DataFrame, DataFrame]): Dataframes

    """
    estoque, df_materiais = setup_data
    resultado = qtd_max("A", estoque, 100, df_materiais)
    esperado = pd.DataFrame(
        {
            "Material": ["A"],
            "Quantidade": [150],
        },
        index=[1],
    )
    pd.testing.assert_frame_equal(resultado, esperado)


def test_material_no_estoque_sem_resultado(setup_data: tuple[DataFrame, DataFrame]) -> None:
    """Testa se a quantidade máxima é retornada corretamente.

    Args:
    ----
        setup_data (tuple[DataFrame, DataFrame]): Dataframes

    """
    estoque, df_materiais = setup_data
    resultado = qtd_max("A", estoque, 200, df_materiais)
    esperado = pd.DataFrame(columns=["Material", "Quantidade"]).astype({"Material": object, "Quantidade": int})
    pd.testing.assert_frame_equal(resultado, esperado)


def test_material_nao_no_estoque(setup_data: tuple[DataFrame, DataFrame]) -> None:
    """Testa se a quantidade máxima é retornada corretamente.

    Args:
    ----
        setup_data (tuple[DataFrame, DataFrame]): Dataframes

    """
    estoque, df_materiais = setup_data
    resultado = qtd_max("D", estoque, 100, df_materiais)
    esperado = pd.DataFrame(columns=["Material", "Quantidade"]).astype({"Material": object, "Quantidade": int})
    pd.testing.assert_frame_equal(resultado, esperado)


if __name__ == "__main__":
    pytest.main()
