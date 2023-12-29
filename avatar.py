# avatar.py
'''Módulo profile da Val.'''
import numpy as np
from PIL import Image


def val_avatar(caminho):
    '''Exibe a imagem da Val no terminal.

    Parâmetros:
    caminho (str): O caminho para a imagem.

    Saída:
    None
    '''

    imagem = Image.open(caminho)
    # pylint: disable=W0612
    imagem_array = np.array(imagem)

    # Redimensiona a imagem para caber no terminal
    altura_terminal, largura_terminal = 30, 80
    imagem_redimensionada = imagem.resize((largura_terminal, altura_terminal))

    # Converte a imagem redimensionada para escala de cinza
    imagem_redimensionada = imagem_redimensionada.convert('L')

    # Exibe a imagem no terminal
    for linha in range(altura_terminal):
        for coluna in range(largura_terminal):
            pixel = imagem_redimensionada.getpixel((coluna, linha))
            caractere = '*' if pixel < 128 else ' '
            print(caractere, end='')
        print()
