# ServicosExecutados.py
"""Módulo de TSE."""

from __future__ import annotations

import typing

import numpy as np

from python.src.excel_tbs import load_worksheets
from python.src.lista_reposicao import dict_reposicao
from python.src.tsepai import pai_dicionario

if typing.TYPE_CHECKING:
    from win32com.client import CDispatch
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
    tb_tse_reposicao,
    tb_tse_retrabalho,
) = load_worksheets()


def verifica_tse(
    servico: CDispatch,
    contrato: str,
    session: CDispatch,
) -> tuple[
    list[str],
    int,
    str | None,
    list[str],
    bool,
    list[tuple[str, str, str, list[str], list[str]]],
    list[tuple[str, str, str, list[str], list[str]]],
    tuple[str, str, str, list[str], list[str]] | None,
    tuple[str, str, str, list[str], list[str]] | None,
    np.ndarray | None,
]:
    """Agrupador de serviço e indexador de classes."""
    empresa, *_ = contrato
    sondagem = ["591000", "567000", "321000", "321500", "283000", "283500"]
    desobstrucao = [
        "561000",
        "563000",
        "568000",
        "581000",
        "584000",
        "585000",
        "586000",
        "587000",
        "592000",
        "717000",
    ]
    troca_pe_cv_prev = ["153000", "153500"]
    pai_tse = 0
    # Lista temporária para armazenar as tse
    list_chave_rb_despesa: list[tuple[str, str, str, list[str], list[str]]] = []
    list_chave_unitario: list[tuple[str, str, str, list[str], list[str]]] = []
    tse_temp: list[str] = []
    identificador_list: list[str] = []
    num_tse_linhas: int = servico.RowCount
    rem_base_reposicao: list[str] = []
    unitario_reposicao: list[str] = []
    mae: bool = False
    tse_proibida = None
    chave_unitario: tuple[str, str, str, list[str], list[str]] | None = None
    chave_rb_despesa: tuple[str, str, str, list[str], list[str]] | None = None
    chave_rb_investimento: tuple[str, str, str, list[str], list[str]] | None = None
    for n_tse in range(num_tse_linhas):
        sap_tse: str = servico.GetCellValue(n_tse, "TSE")
        etapa_pai: str = servico.GetCellValue(n_tse, "ETAPA")

        # Pulando OS com asfalto incluso.
        if sap_tse in dict_reposicao["asfalto"]:
            tse_proibida = "Aslfato na bagaça!"
            break

        # Reposição de Guia, fazer manual.
        if sap_tse == "755000":
            tse_proibida = "Reposicao de Guia!"
            break

        # Retirado Entulho.
        if sap_tse == "714000":
            tse_proibida = "Retirado Entulho"
            break

        # REPOSIÇÃO DE PARALELO , fazer manual.
        if sap_tse == "782500":
            tse_proibida = "REPOSIÇÃO DE SARJETA INV"
            reposicao = sap_tse
            etapa_reposicao = etapa_pai
            break

        # REGULARIZADO RAMAL DE AGUA DESVIADO
        if sap_tse == "282000":
            tse_proibida = "REGULARIZADO RAMAL DE AGUA DESVIADO"
            break

        # Serviços relacionados a obra.
        if sap_tse in (
            "300000",
            "308000",
            "310000",
            "311000",
            "313000",
            "315000",
            "532000",
            "564000",
            "588000",
            "590000",
            "709000",
            "700000",
            "593000",
        ):
            tse_proibida = "Obra."
            break

        # RAMAL AGUA UNITÁRIO - A PEDIDO DA IARA
        if sap_tse in (
            "253000",
            "250000",
            "209000",
            "605000",
            "605000",
            "263000",
            "255000",
            "254000",
            "282000",
            "265000",
            "260000",
            "265000",
            "263000",
            "262000",
            "284500",
            "286000",
            "282500",
        ):
            tse_proibida = "Iara não quer."
            break

        """---------- REVOGADO --------------------------------
        # TROCA PÉ DE CV PREVENTIVO - Solicitado por Ivan/Estevan
        # if sap_tse in ('153000', '153500'):
        #     tse_proibida = 'Ivan não quer.'
        #     break
        """

        # INSTALDO CAIXA UMA (PARTE CIVIL)
        if sap_tse == "136000":
            tse_proibida = "Instalado Caixa Uma."
            break

        # TROCA DE CAVALETE POR UMA E RELIGAÇÃO
        if sap_tse == "159000":
            tse_proibida = "TROCA POR UMA E RELIGADA"
            break

        # SUBSTITUIDA TAMPA DE CAIXA UMA
        # Não definido o que fazer.
        if sap_tse == "155000":
            tse_proibida = "SUBSTITUIDA TAMPA DE CAIXA UMA"
            break

        if sap_tse in tb_tse_un:  # Verifica se está no Conjunto Unitários
            servico.modifyCell(n_tse, "PAGAR", "s")  # Marca pagar na TSE
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            (reposicao, tse_proibida, identificador, etapa_reposicao) = pai_dicionario.pai_servico_unitario(sap_tse, session)
            identificador_list.append(identificador)
            chave_unitario = (
                sap_tse,
                etapa_pai,
                identificador,
                reposicao,
                etapa_reposicao,
            )
            unitario_reposicao.extend(reposicao)
            list_chave_unitario.append(chave_unitario)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        if sap_tse in sondagem or sap_tse in tb_tse_rem_base:  # Caso Contrário, é RB - Sondagem
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5")  # Despesa
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            (reposicao, tse_proibida, identificador, etapa_reposicao) = pai_dicionario.pai_servico_cesta(sap_tse, session)
            rem_base_reposicao.extend(reposicao)
            identificador_list.append(identificador)
            chave_rb_despesa = (
                sap_tse,
                etapa_pai,
                identificador,
                reposicao,
                etapa_reposicao,
            )
            list_chave_rb_despesa.append(chave_rb_despesa)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        if sap_tse in tb_tse_invest:  # Caso Contrário, é RB - Investimento
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "6")  # Investimento
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            (reposicao, tse_proibida, identificador, etapa_reposicao) = pai_dicionario.pai_servico_cesta(sap_tse, session)
            rem_base_reposicao.extend(reposicao)
            identificador_list.append(identificador)
            chave_rb_investimento = (
                sap_tse,
                etapa_pai,
                identificador,
                reposicao,
                etapa_reposicao,
            )
            pai_tse += 1
            mae = True
            if tse_proibida is not None:
                break
            continue

        if sap_tse in desobstrucao and empresa == "4600043760":
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "5")  # Despesa
            # Coloca a tse existente na lista temporária
            tse_temp.append(sap_tse)
            # pylint: disable=E1121
            (reposicao, tse_proibida, identificador, etapa_reposicao) = pai_dicionario.pai_servico_desobstrucao(
                sap_tse,
                session,
            )
            identificador_list.append(identificador)
            chave_rb_despesa = (
                sap_tse,
                etapa_pai,
                identificador,
                reposicao,
                etapa_reposicao,
            )
            list_chave_rb_despesa.append(chave_rb_despesa)
            pai_tse += 1
            if tse_proibida is not None:
                break
            continue

        if sap_tse in tb_tse_nexec:
            servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
            servico.modifyCell(n_tse, "CODIGO", "11")  # Não Executado
            continue

        if sap_tse in tb_tse_pertence_ao_servico_principal:
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Pertence ao Serviço Principal
            servico.modifyCell(n_tse, "CODIGO", "3")
            continue

        if sap_tse in tb_tse_retrabalho:
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "7")  # Retrabalho
            continue

        if sap_tse in ("730600", "730700"):
            # Compactação e Selagem da Base.
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Pedido alterado por Iara Regina
            servico.modifyCell(n_tse, "CODIGO", "5")
            continue

        if sap_tse in ("761000", "762000"):
            # REPOSIÇÃO DE PAREDE/MURO INV.
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Pedido alterado por Iara Regina
            servico.modifyCell(n_tse, "CODIGO", "5")
            continue

        if sap_tse == "666000":
            # VISTORIADO LOCAL E IDENTIFICADA SITUAÇÃO
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "10")  # Serviço MOP
            continue

        if sap_tse == "408000":
            # SUPRIMIDA LIG POÇO ENCERRAMENTO CONTRATO
            servico.modifyCell(n_tse, "PAGAR", "n")
            servico.modifyCell(n_tse, "CODIGO", "10")  # Serviço MOP
            continue

        if sap_tse in desobstrucao and empresa != "4600043760":
            # FUMAÇA, DD/DC, LAVAGEM, TELEVISIONADO
            servico.modifyCell(n_tse, "PAGAR", "n")
            # Serviço não existe no contrato
            servico.modifyCell(n_tse, "CODIGO", "10")
            continue

        if empresa == "4600043760" and sap_tse not in desobstrucao:
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
            reposicao_geral,
        )

    # Manipulação das matrizes
    rem_base_reposicao_union = np.unique(rem_base_reposicao, axis=None)
    unitario_reposicao_flat = np.ravel(unitario_reposicao)
    reposicao_geral = np.unique(
        np.concatenate([rem_base_reposicao_union, unitario_reposicao_flat]),
    )

    # TSEs situacionais.
    if chave_unitario is not None and pai_tse == 1 and chave_unitario[0] in troca_pe_cv_prev:
        for n_tse in range(num_tse_linhas):
            sap_tse = servico.GetCellValue(n_tse, "TSE")
            etapa_pai = servico.GetCellValue(n_tse, "ETAPA")
            # Altera todas as reposições para N3 de Troca Pé Preventivo.
            if sap_tse in rem_base_reposicao_union:
                servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço Principal
                servico.modifyCell(n_tse, "CODIGO", "3")

    if (
        chave_rb_despesa is not None
        and pai_tse == 1
        or (chave_rb_despesa is not None and all(tse in sondagem for tse in chave_rb_despesa[0]))
    ) and chave_rb_despesa[0] in sondagem:
        for n_tse in range(num_tse_linhas):
            sap_tse = servico.GetCellValue(n_tse, "TSE")
            etapa_pai = servico.GetCellValue(n_tse, "ETAPA")
            # Altera todas as reposições de rb para N3 de Sondagem.
            # TODO (Hyslan): Alterar para N5 e itens não vinculados.
            if sap_tse in rem_base_reposicao_union:
                servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                # Pertence ao serviço Principal
                servico.modifyCell(n_tse, "CODIGO", "3")

    if mae is True:
        for n_tse in range(num_tse_linhas):
            sap_tse = servico.GetCellValue(n_tse, "TSE")
            etapa_pai = servico.GetCellValue(n_tse, "ETAPA")
            # Altera todas as reposições de rb para investimento se tiver tra.
            if sap_tse in rem_base_reposicao_union:
                servico.modifyCell(n_tse, "PAGAR", "n")  # Cesta
                servico.modifyCell(n_tse, "CODIGO", "6")  # Investimento
    # Fim da condicional.
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
        reposicao_geral,
    )
