# temporizador.py
'''Módulo Cronomêtro.'''
import time


def cronometro_val(start_time, ordem):
    '''Cronomêtro da Val.'''
    end_time = time.time()
    # Tempo de execução.
    execution_time = end_time - start_time
    print(f"Tempo gasto para valorar a Ordem: {ordem}, "
          + f"foi de: {execution_time} segundos.")
    print(f"****Fim da Valoração da Ordem: {ordem} ****")
