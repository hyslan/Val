# temporizador.py
"""Módulo Cronomêtro."""
import time

from rich.console import Console
from rich.panel import Panel

console = Console()


def cronometro_val(start_time: float, ordem: str) -> float:
    """Cronomêtro da Val."""
    end_time = time.time()
    # Tempo de execução.
    execution_time = end_time - start_time
    console.print(Panel.fit(f"Tempo gasto para valorar a Ordem: {ordem}, "
                            + f"foi de: {round(execution_time, 2)} segundos."), style="italic cyan")
    console.print(Panel.fit(f"Fim da Valoração da Ordem: {
                  ordem}"), style="bold green")
    return round(execution_time, 2)
