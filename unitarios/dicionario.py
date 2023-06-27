# dicionario.py
'''Módulo do Dicionário de Unitários'''
# Bibliotecas
import sys
from sap_connection import connect_to_sap
from unitarios.hidrometro import m_hidrometro
from unitarios.supressao import m_supressao
from unitarios.religacao import m_religacao
from unitarios.cavalete import m_cavalete

session = connect_to_sap()


def preserv_inter_serv(identificador):
    '''Etapa de preservação de Interferência'''
    ramo_agua = [
        ["hidrometro"],
        ["cavalete"],
        ["religacao"],
        ["supressao"],
        ["ramal_agua"],
        ["tra"],
        ["ligacao_agua"],
        ["rede_agua"],
        ["supr_restab"]
    ]
    identificador.remove("preservacao")
    if identificador in ramo_agua:
        serv_preservacao = '456118'
    else:
        serv_preservacao = '456217'

    preco = session.findById(
        "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
        + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
    preco.GetCellValue(0, "NUMERO_EXT")
    if preco is not None:
        num_precos_linhas = preco.RowCount
        print(
            f"Qtd linhas em itens de preço: {num_precos_linhas}")
        n_preco = 0  # índice para itens de preço
        for n_preco, sap_preco in enumerate(range(0, num_precos_linhas)):
            sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
            if sap_preco == serv_preservacao:
                # Marca pagar na TSE
                preco.modifyCell(n_preco, "QUANT", "1")
                preco.setCurrentCell(n_preco, "QUANT")
                preco.pressEnter()
                print(
                    f"Pago 1 UN de Preservação - CODIGO: {serv_preservacao}")
                break


def unitario(etapa, corte, relig, reposicao, num_tse_linhas, etapa_reposicao, identificador):
    '''Dicionário de chaves para etapas de unitário.'''
    dicionario_un = {
        '201000': m_hidrometro.Hidrometro.THD_456901,
        '203000': m_hidrometro.Hidrometro.THD_456901,
        '203500': m_hidrometro.Hidrometro.HD_456022,
        '204000': m_hidrometro.Hidrometro.THD_456901,
        '205000': m_hidrometro.Hidrometro.THD_456901,
        '206000': m_hidrometro.Hidrometro.THD_456901,
        '207000': m_hidrometro.Hidrometro.THD_456901,
        '215000': m_hidrometro.Hidrometro.THDPrev_456902,
        '405000': m_supressao.Corte.supressao,
        '414000': m_supressao.Corte.supressao,
        '450500': m_religacao.Religacao.restabelecida,
        '453000': m_religacao.Religacao.restabelecida,
        '455500': m_religacao.Religacao.restabelecida,
        '463000': m_religacao.Religacao.restabelecida,
        '465000': m_religacao.Religacao.restabelecida,
        '467500': m_religacao.Religacao.restabelecida,
        '475500': m_religacao.Religacao.restabelecida,
        # '142000': m_cavalete.Cavalete.
        # '148000': m_cavalete.Cavalete.
        '149000': m_cavalete.Cavalete.TrocaCvKit,
        '153000': m_cavalete.Cavalete.TrocaPeCvPrev,

    }

    if etapa in ('713000', '713500'):
        preserv_inter_serv(identificador)
    else:
        if etapa in dicionario_un:
            print(f"Etapa está inclusa no Dicionário de Unitários: {etapa}")
            metodo = dicionario_un[etapa]
            # Chama o método de uma classe dentro do Dicionário
            metodo(corte, relig, reposicao, num_tse_linhas, etapa_reposicao)
        else:
            print("TSE não Encontrada no Dicionário!")
            sys.exit()
