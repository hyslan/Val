#cesta_dicionario.py
'''Dicionário da Remuneração Base do Global.'''
#Bibliotecas
import sys

def cesta(etapa, corte, relig):
    '''Condicional do dicionário.'''
    dicionario_rb = {
        #'130000': m_hidrometro.Hidrometro.THD_456901,

        # Adicionar chaves conforme classes e métodos forem adicionados.
    }

    if etapa in dicionario_rb:
        print(f"Etapa está inclusa no Dicionário da Cesta: {etapa}")
        metodo = dicionario_rb[etapa]
        metodo(corte, relig)  # Chama o método de uma classe dentro do Dicionário
    else:
        print("TSE não Encontrada no Dicionário!")
        sys.exit()
