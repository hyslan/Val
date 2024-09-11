# interfaces.py
"""Establish a pattern."""

from abc import ABC, abstractmethod


class UnitarioInterface(ABC):
    """Interface para as classes."""

    @abstractmethod
    def processar(self) -> None:
        """Processar tse."""

    @abstractmethod
    def pagar(self) -> None:
        """Abstrato para pagar."""

    @abstractmethod
    def _processar_operacao(self, tipo_operacao: str) -> None:
        """Processar operação."""
