# salvacao.py
'''Módulo de salvar valoração.'''
import threading
import pywintypes
import sql_view
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from confere_os import consulta_os
from sap import fechar_conexao

(
    lista,
    _,
    _,
    _,
    planilha,
    _,
    _,
    _,
    _,
    _,
    _,
    _,
    *_,
) = load_worksheets()


def salvar(ordem, qtd_ordem):
    '''Salvar e verificar se está salvando.'''

    def salvar_valoracao():
        '''Função para salvar valoração.'''
        nonlocal ordem
        nonlocal qtd_ordem
        session = connect_to_sap()
        try:
            print("Salvando valoração!")
            session.findById("wnd[0]").sendVKey(11)
            session.findById("wnd[1]/usr/btnBUTTON_1").press()
        # pylint: disable=E1101
        except pywintypes.com_error:
            print(f"Ordem: {ordem} não foi salva.")
            ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
            ja_valorado.valorada(obs="Não foi salvo")

        # Verificar se Salvou
        (status_sistema,
            status_usuario,
            *_) = consulta_os(ordem)
        print("Verificando se Ordem foi valorada.")
        if status_usuario == "EXEC VALO":
            print(f"Status da Ordem: {status_sistema}, {status_usuario}")
            print("Foi Salvo com sucesso!")
            ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
            ja_valorado.valorada("SIM")
            # Incremento + de Ordem.
            qtd_ordem += 1
        else:
            print(f"Ordem: {ordem} não foi salva.")
            ja_valorado = sql_view.Tabela(ordem=ordem, cod_tse="")
            ja_valorado.valorada(obs="Não foi salvo")

    # Thread para função aninhada.
    thread_salvar = threading.Thread(target=salvar_valoracao)
    thread_salvar.start()
    # Aguardar até o limite de 5min.
    thread_salvar.join(timeout=300)
    # Verificar se está em execução.
    if thread_salvar.is_alive():
        print("SAP demorando mais que o esperado, encerrando.")
        fechar_conexao()

    return qtd_ordem
