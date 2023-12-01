# dicionario.py
'''Módulo do Dicionário de Unitários'''
# Bibliotecas
import sys
from sap_connection import connect_to_sap
from unitarios.hidrometro import m_hidrometro
from unitarios.supressao import m_supressao
from unitarios.religacao import m_religacao
from unitarios.cavalete import m_cavalete
from unitarios.poco import m_poco
from unitarios.ligacao_esgoto import m_ligacao_esgoto
from unitarios.ligacao_agua import m_ligacao_agua


def preserv_inter_serv(identificador):
    '''Etapa de preservação de Interferência'''
    session = connect_to_sap()
    ramo_agua = [
        ["hidrometro"],
        ["cavalete"],
        ["religacao"],
        ["supressao"],
        ["ramal_agua"],
        ["tra"],
        ["ligacao_agua"],
        ["ligacao_agua_nova"],
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


def unitario(etapa,
             corte,
             relig,
             reposicao,
             num_tse_linhas,
             etapa_reposicao,
             identificador,
             posicao_rede,
             profundidade
             ):
    '''Dicionário de chaves para etapas de unitário.'''
    dicionario_un = {
        '134000': m_cavalete.Cavalete.instalado_lacre,
        '135000': m_cavalete.Cavalete.instalado_lacre,
        # '142000': m_cavalete.Cavalete.
        '148000': m_cavalete.Cavalete.troca_cv_por_uma,
        '149000': m_cavalete.Cavalete.troca_cv_kit,
        '153000': m_cavalete.Cavalete.troca_pe_cv_prev,
        '153500': m_cavalete.Cavalete.troca_pe_cv_prev,
        '201000': m_hidrometro.Hidrometro.troca_de_hidro_corretivo,
        '202000': m_hidrometro.Hidrometro.desinclinado_hidrometro,
        '203000': m_hidrometro.Hidrometro.troca_de_hidro_corretivo,
        '203500': m_hidrometro.Hidrometro.desinclinado_hidrometro,
        '204000': m_hidrometro.Hidrometro.troca_de_hidro_corretivo,
        '205000': m_hidrometro.Hidrometro.troca_de_hidro_corretivo,
        '206000': m_hidrometro.Hidrometro.troca_de_hidro_corretivo,
        '207000': m_hidrometro.Hidrometro.troca_de_hidro_corretivo,
        '215000': m_hidrometro.Hidrometro.troca_de_hidro_preventiva_agendada,
        '254000': m_ligacao_agua.LigacaoAgua.ligacao_agua,
        '280000': m_ligacao_agua.LigacaoAgua.png,
        '284500': m_ligacao_agua.LigacaoAgua.tra_nv,
        '312000': m_poco.Poco.niv_cx_parada,
        '322000': m_poco.Poco.troca_de_caixa_de_parada,
        '405000': m_supressao.Corte.supressao,
        '406000': m_supressao.Corte.supressao,
        '407000': m_supressao.Corte.supressao,
        '414000': m_supressao.Corte.supressao,
        '450500': m_religacao.Religacao.restabelecida,
        '453000': m_religacao.Religacao.restabelecida,
        '455500': m_religacao.Religacao.restabelecida,
        '463000': m_religacao.Religacao.restabelecida,
        '465000': m_religacao.Religacao.restabelecida,
        '467500': m_religacao.Religacao.restabelecida,
        '475500': m_religacao.Religacao.restabelecida,
        '502000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
        '505000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
        '506000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
        '508000': m_ligacao_esgoto.LigacaoEsgoto.ligacao_esgoto,
        '537000': m_poco.Poco.nivelamento,
        '537100': m_poco.Poco.nivelamento,
        '538000': m_poco.Poco.nivelamento,
        '565000': m_ligacao_esgoto.LigacaoEsgoto.png_esgoto,
        '569000': m_ligacao_esgoto.LigacaoEsgoto.tre

    }

    if etapa in ('713000', '713500'):
        preserv_inter_serv(identificador)
    else:
        if etapa in dicionario_un:
            print(f"Etapa está inclusa no Dicionário de Unitários: {etapa}")
            metodo = dicionario_un[etapa]
            # Chama o método de uma classe dentro do Dicionário
            metodo(corte,
                   relig,
                   reposicao,
                   num_tse_linhas,
                   etapa_reposicao,
                   posicao_rede,
                   profundidade
                   )
        else:
            print("TSE não Encontrada no Dicionário!")
            sys.exit()
