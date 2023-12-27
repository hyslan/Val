from rich.console import Console
from confere_os import consulta_os
console = Console()
def main():
    try:
        consulta_os("", "4600041302", "344")
    except Exception as erro:
        console.print(f"Erro: {erro}")
        console.print_exception(show_locals=True)

# Executar o loop de eventos asyncio
if __name__ == "__main__":
    main()
