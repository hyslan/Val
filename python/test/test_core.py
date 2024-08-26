"""Módulo de testes do módulo core."""

import pandas as pd

from src.core import val

pendentest_array = pd.DataFrame(["2425188908", "100", "4600041302"], index=[0], columns=["Ordem", "COD_MUNICIPIO", "Contrato"])


def test_core_val():
    val()
    # TODO: Implementar testes
