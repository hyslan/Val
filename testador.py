# testador.py
'''Área de testes.'''
import subprocess

# Substitua 'nome_do_processo' pelo nome real do processo que você deseja encerrar
processo_a_encerrar = 'saplogon.exe'

# Tenta encerrar o processo
try:
    subprocess.run(['taskkill', '/F', '/IM', processo_a_encerrar], check=True)
    print(f'O processo {processo_a_encerrar} foi encerrado com sucesso.')
except subprocess.CalledProcessError:
    print(f'Não foi possível encerrar o processo {processo_a_encerrar}.')
