# salvacao.py
'''Módulo de salvar valoração.'''
import sys
import asyncio
import pywintypes
import sql_view
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from confere_os import consulta_os
from sap import encerrar_sap

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


async def salvar(ordem, qtd_ordem):
    '''Salvar e verificar se está salvando.'''

    async def salvar_valoracao():
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

    # função aninhada.
    try:
        # Timeout = 5min
        await asyncio.wait_for(salvar_valoracao(), timeout=300)
    except asyncio.TimeoutError:
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()

    return qtd_ordem
