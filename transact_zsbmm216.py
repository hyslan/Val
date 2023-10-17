# transact_zsbmm216.py
'''Módulo de Contrato'''
# Conexão SAP
import asyncio
from sap_connection import connect_to_sap
from sap import encerrar_sap


async def novasp(ordem):
    '''Async NOVAS SP'''

    async def t_novasp():
        '''Transação preenchida ZSBMM216 - Contrato NOVASP'''
        nonlocal ordem
        session = connect_to_sap()
        print("Iniciando valoração.")
        session.StartTransaction("ZSBMM216")
        # Unidade Administrativa
        session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
        # Contrato NOVASP
        session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600041302"
        session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
        sap_ordem = session.findById("wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
        sap_ordem.Text = ordem
        session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # função aninhada.
    try:
        # Timeout = 5min
        await asyncio.wait_for(t_novasp(), timeout=300)
    except asyncio.TimeoutError:
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()


async def recape(ordem):
    '''Async RECAPE'''

    async def t_recape():
        '''Transação preenchida ZSBMM216 - Contrato RECAPE'''
        nonlocal ordem
        session = connect_to_sap()
        print("Iniciando valoração.")
        session.StartTransaction("ZSBMM216")
        # Unidade Administrativa
        session.findById("wnd[0]/usr/ctxtP_UND").Text = "344"
        # Contrato RECAPE
        session.findById("wnd[0]/usr/ctxtP_CONT").Text = "4600044782"
        session.findById("wnd[0]/usr/ctxtP_MUNI").Text = "100"  # Município
        sap_ordem = session.findById("wnd[0]/usr/ctxtP_ORDEM")  # Campo ordem
        sap_ordem.Text = ordem
        session.findById("wnd[0]").SendVkey(8)  # Aperta botão F8

    # função aninhada.
    try:
        # Timeout = 5min
        await asyncio.wait_for(t_recape(), timeout=300)
    except asyncio.TimeoutError:
        print("SAP demorando mais que o esperado, encerrando.")
        encerrar_sap()
