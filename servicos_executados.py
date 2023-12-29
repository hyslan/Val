# ServicosExecutados.py
'''Módulo de TSE'''
import numpy as np
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from tsepai import pai_dicionario


(
    lista,
    materiais,
    plan_tse,
    plan_precos,
    planilha,
    contratada,
    unitario,
    rem_base,
    naoexecutado,
    invest,
    tb_contratada,
    tb_contratada_gb,
    tb_tse_un,
    tb_tse_rem_base,
    tb_tse_nexec,
    tb_tse_invest,
    tb_st_sistema,
    tb_st_usuario,
    tb_tse_pertence_ao_servico_principal,
    tb_tse_servico_nao_existe_no_contrato,
    tb_tse_reposicao,
    tb_tse_retrabalho,
    tb_tse_asfalto,
) = load_worksheets()


def verifica_tse(servico, contrato):
    '''Agrupador de serviço e indexador de classes.'''
    session = connect_to_sap()
    sondagem = [
        '591000',
        '567000',
        '321000',
        '321500',
        '283000',
        '283500'
    ]
    desobstrucao = [
        '561000',
        '563000',
        '568000',
        '581000',
        '584000',
        '585000',
        '586000',
        '587000',
        '592000',
        '717000'
    ]
    troca_pe_cv_prev = ['153000', '153500']
    pai_tse = 0
    print("Iniciando processo de verificação de TSE")
    servico = session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABS/ssubSUB_TAB:"
                               + "ZSBMM_VALORACAOINV:9010/cntlCC_SERVICO/shellcont/shell")
    # Lista temporária para armazenar as tse
    list_chave_rb_despesa = []
    list_chave_unitario = []
    tse_temp = []
    identificador_list = []
    num_tse_linhas = servico.RowCount
    rem_base_reposicao = []
    unitario_reposicao = []
    mae = False
    chave_unitario = None
    chave_rb_despesa = None
    chave_rb_investimento = None
    print(f"Qtd de linhas de serviços executados: {num_tse_linhas}")
    for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
        sap_tse = servico.GetCellValue(n_tse, "TSE")
        etapa_pai = servico.GetCellValue(n_tse, "ETAPA")

        if sap_tse in tb_tse_un:  # Verifica se está no Conjunto Unitários
            servico.modifyCell(n_tse, "PAGAR", "s")  # Marca pagar na TSE
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            # pylint: disable=E1121
            (reposicao,
             tse_proibida,
             identificador,
             etapa_reposicao) = pai_dicionario.pai_servico_unitario(sap_tse)
            identificador_list.append(identificador)
            chave_unitario = sap_tse, etapa_pai, identificador, reposicao, etapa_reposicao
            unitario_reposicao.append(reposicao)
            list_chave_unitario.append(chave_unitario)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        elif sap_tse in sondagem:  # Caso Contrário, é RB - Sondagem
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5")  # Despesa
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            (reposicao,
             tse_proibida,
             identificador,
             etapa_reposicao) = pai_dicionario.pai_servico_cesta(sap_tse)
            rem_base_reposicao.append(reposicao)
            identificador_list.append(identificador)
            chave_rb_despesa = sap_tse, etapa_pai, identificador, reposicao, etapa_reposicao
            list_chave_rb_despesa.append(chave_rb_despesa)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        elif sap_tse in tb_tse_rem_base:  # Caso Contrário, é RB - Despesa
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5")  # Despesa
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            (reposicao,
             tse_proibida,
             identificador,
             etapa_reposicao) = pai_dicionario.pai_servico_cesta(sap_tse)
            rem_base_reposicao.append(reposicao)
            identificador_list.append(identificador)
            chave_rb_despesa = sap_tse, etapa_pai, identificador, reposicao, etapa_reposicao
            list_chave_rb_despesa.append(chave_rb_despesa)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        elif sap_tse in tb_tse_invest:  # Caso Contrário, é RB - Investimento
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "6")  # Investimento
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            (reposicao,
             tse_proibida,
             identificador,
             etapa_reposicao) = pai_dicionario.pai_servico_cesta(sap_tse)
            rem_base_reposicao.append(reposicao)
            identificador_list.append(identificador)
            chave_rb_investimento = sap_tse, etapa_pai, identificador, reposicao, etapa_reposicao
            pai_tse += 1
            mae = True
            if tse_proibida is not None:
                break
            continue

        elif sap_tse in desobstrucao and contrato == "4600043760":
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5")  # Despesa
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            # pylint: disable=E1121
            (reposicao,
             tse_proibida,
             identificador,
             etapa_reposicao) = pai_dicionario.pai_servico_desobstrucao(sap_tse)
            identificador_list.append(identificador)
            chave_rb_despesa = sap_tse, etapa_pai, identificador, reposicao, etapa_reposicao
            list_chave_rb_despesa.append(chave_rb_despesa)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        # Pulando OS com asfalto incluso.
        elif sap_tse in tb_tse_asfalto:
            tse_proibida = "Aslfato na bagaça!"
            break

        # Reposição de Guia, fazer manual.
        elif sap_tse == '755000':
            tse_proibida = 'Reposicao de Guia!'
            break

        # Retirado Entulho.
        elif sap_tse == '714000':
            tse_proibida = 'Retirado Entulho'
            break

        # REPOSIÇÃO DE PARALELO , fazer manual.
        elif sap_tse == '782500':
            tse_proibida = "REPOSIÇÃO DE SARJETA INV"
            reposicao = sap_tse
            etapa_reposicao = etapa_pai
            break

        # Compactação e selagem da base.
        elif sap_tse in ('758500', '758000'):
            tse_proibida = 'PARALELO'
            break

        # Readequado Cavalete, verificar...
        # elif sap_tse == '138000':
        #     tse_proibida = 'Readequado Cavalete!'
        #     break

        # Suprimido Ramal anterior
        elif sap_tse == '415000':
            tse_proibida = 'Ramal anterior'
            break

        # SUPRIMIDO RAMAL AGUA ABAND NÃO VISIVEL
        elif sap_tse == '416500':
            tse_proibida = "Verificar."
            break

        # Serviços relacionados a obra.
        elif sap_tse in ('300000', '308000', '310000', '311000', '313000',
                         '315000', '532000', '564000', '588000', '590000',
                         '709000', '700000', '593000'):
            tse_proibida = 'Obra.'
            break

        elif sap_tse in tb_tse_nexec:
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "11")  # Não Executado
            continue

        elif sap_tse in tb_tse_pertence_ao_servico_principal:
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Pertence ao Serviço Principal
            servico.modifyCell(n_tse, "CODIGO", "3")
            continue

        elif sap_tse in tb_tse_retrabalho:
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "7")  # Retrabalho
            continue

        elif sap_tse in ('730600', '730700'):
            # Compactação e Selagem da Base.
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "1")  # Divergência
            continue

        elif sap_tse in ('761000', '762000'):
            # REPOSIÇÃO DE PAREDE/MURO INV.
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "1")  # Divergência
            continue

        elif sap_tse == '666000':
            # VISTORIADO LOCAL E IDENTIFICADA SITUAÇÃO
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "10")  # Serviço MOP
            continue

        if sap_tse in desobstrucao and not contrato == "4600043760":
            # FUMAÇA, DD/DC, LAVAGEM, TELEVISIONADO
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Serviço não existe no contrato
            servico.modifyCell(n_tse, "CODIGO", "10")
            continue

        if contrato == "4600043760" and sap_tse not in desobstrucao:
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "10")
            continue

    if tse_proibida is not None:
        reposicao = None
        etapa_reposicao = None
        rem_base_reposicao_union = None
        reposicao_geral = None
        return (
            tse_temp,
            num_tse_linhas,
            tse_proibida,
            identificador_list,
            mae,
            list_chave_rb_despesa,
            list_chave_unitario,
            chave_rb_investimento,
            chave_unitario,
            reposicao_geral
        )

    # Manipulação das matrizes
    rem_base_reposicao_union = np.unique(rem_base_reposicao, axis=None)
    unitario_reposicao_flat = np.ravel(unitario_reposicao)
    reposicao_geral = np.unique(np.concatenate(
        [rem_base_reposicao_union, unitario_reposicao_flat]))

    # TSEs situacionais.
    if chave_unitario is not None and pai_tse == 1:
        if chave_unitario[0] in troca_pe_cv_prev:
            for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
                sap_tse = servico.GetCellValue(n_tse, "TSE")
                etapa_pai = servico.GetCellValue(n_tse, "ETAPA")
                # Altera todas as reposições para N3 de Troca Pé Preventivo.
                if sap_tse in rem_base_reposicao_union:
                    servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                    # Pertence ao serviço Principal
                    servico.modifyCell(n_tse, "CODIGO", "3")

    if chave_rb_despesa is not None and pai_tse == 1 \
            or all(tse in sondagem for tse in chave_rb_despesa[0]):
        if chave_rb_despesa[0] in sondagem:
            for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
                sap_tse = servico.GetCellValue(n_tse, "TSE")
                etapa_pai = servico.GetCellValue(n_tse, "ETAPA")
                # Altera todas as reposições de rb para N3 de Sondagem.
                if sap_tse in rem_base_reposicao_union:
                    servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                    # Pertence ao serviço Principal
                    servico.modifyCell(n_tse, "CODIGO", "3")

    if mae is True:
        for n_tse, sap_tse in enumerate(range(0, num_tse_linhas)):
            sap_tse = servico.GetCellValue(n_tse, "TSE")
            etapa_pai = servico.GetCellValue(n_tse, "ETAPA")
            # Altera todas as reposições de rb para investimento se tiver tra.
            if sap_tse in rem_base_reposicao_union:
                servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                servico.modifyCell(n_tse, "CODIGO", "6")  # Investimento
    # Fim da condicional.
    # sys.exit()
    servico.pressEnter()

    return (
        tse_temp,
        num_tse_linhas,
        tse_proibida,
        identificador_list,
        mae,
        list_chave_rb_despesa,
        list_chave_unitario,
        chave_rb_investimento,
        chave_unitario,
        reposicao_geral
    )
