import time
from rich.console import Console

console = Console()

with console.status("[bold blue]Trabalhando..."):
    for i in range(100):
        print(f"Prints: {i}")
        time.sleep(1)
