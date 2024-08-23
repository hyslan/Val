# avatar.py
"""Módulo profile da Val."""

from PIL import Image


def val_avatar() -> None:
    """Exibe a imagem da Val no terminal.

    Parâmetros:
    caminho (str): O caminho para a imagem.

    Saída:
    None
    """
    caminho: str = "media/val.png"
    imagem: Image = Image.open(caminho)

    # Redimensiona a imagem para caber no terminal
    altura_terminal, largura_terminal = 30, 80
    imagem_redimensionada: Image = imagem.resize((largura_terminal, altura_terminal))

    # Converte a imagem redimensionada para escala de cinza
    imagem_redimensionada: Image = imagem_redimensionada.convert("L")

    # Exibe a imagem no terminal
    for linha in range(altura_terminal):
        for coluna in range(largura_terminal):
            imagem_redimensionada.getpixel((coluna, linha))
