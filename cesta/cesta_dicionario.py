#cesta_dicionario.py
import sys


def Cesta(Etapa, Corte, Relig):
    
    DicionarioRB = {
        #'130000': m_hidrometro.Hidrometro.THD_456901,

        # Adicionar chaves conforme classes e métodos forem adicionados ao diretório e instanciados na main
    }

    if Etapa in DicionarioRB:
        print(f"Etapa está inclusa no Dicionário da Cesta: {Etapa}")
        metodo = DicionarioRB[Etapa]
        metodo(Corte, Relig)  # Chama o método de uma classe dentro do Dicionário
    else:
        print("TSE não Encontrada no Dicionário!")
        sys.exit()
