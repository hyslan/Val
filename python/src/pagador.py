# pagador.py
"""Módulo para precificar na aba itens de preço."""
from rich.console import Console
from rich.columns import Columns
from rich.panel import Panel
from python.src.servicos_executados import verifica_tse
from python.src.unitarios.controlador import Controlador
from python.src.cesta import cesta_dicionario


def precificador(tse, corte, relig,
                 posicao_rede, profundidade, contrato,
                 session):
    """Função para apontar os itens de preço e selecionar."""
    console = Console()
    tse.GetCellValue(0, "TSE")  # Saber qual TSE é
    (
        tse_temp,
        num_tse_linhas,
        tse_proibida,
        identificador,
        mae,
        list_chave_rb_despesa,
        list_chave_unitario,
        chave_rb_investimento,
        chave_unitario,
        reposicao_geral
    ) = verifica_tse(
        tse, contrato, session)

    etapa_unitario = []
    ligacao_errada = False
    profundidade_errada = False

    if tse_proibida is not None:
        console.print(
            Columns([Panel("[bold red]TSE proibida de ser valorada.")]), justify="center")
        return (
            tse_proibida,
            list_chave_rb_despesa,
            list_chave_unitario,
            chave_rb_investimento,
            chave_unitario,
            ligacao_errada,
            profundidade_errada
        )

    if chave_unitario is not None and chave_unitario[0] in (
            '253000',
            '254000',
            '255000',
            '262000',
            '263000',
            '265000',
            '266000',
            '268000',
            '269000',
            '280000',
            '284500',
            '286000',
            '502000',
            '505000',
            '508000',
            '561000',
            '565000',
            '569000',
            '581000',
            '539000',
            '539000',
            '585000',
    ) and not posicao_rede:
        ligacao_errada = True

    if chave_unitario is not None and chave_unitario[0] in (
            '502000',
            '505000',
            '508000',
            '561000',
            '565000',
            '569000',
            '581000',
            '539000',
            '539000',
            '585000',
    ) and not profundidade:
        profundidade_errada = True

    if ligacao_errada or profundidade_errada is True:
        print("Sem informação de rede.")
        return (
            tse_proibida,
            list_chave_rb_despesa,
            list_chave_unitario,
            chave_rb_investimento,
            chave_unitario,
            ligacao_errada,
            profundidade_errada
        )

    console.print(Columns([Panel(f"[b]TSE: {tse_temp}")]))
    console.print(
        Columns([Panel(f"Reposição inclusa : {reposicao_geral}")]))
    if list_chave_unitario:
        console.print(
            Columns([Panel(f"[b]Chave unitario: {list_chave_unitario}")]), style="bold magenta")
    if list_chave_rb_despesa or chave_rb_investimento:
        console.print(
            Columns(
                [Panel(f"[b]Chave RB: {list_chave_rb_despesa}, {chave_rb_investimento}")]),
            style="bold magenta")

    if list_chave_unitario:  # Verifica se está no Conjunto Unitários
        # Aba Itens de preço
        session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI").select()

        console.print("Processo de Precificação",
                      style="bold red underline", justify="center")
        for chave_unitario in list_chave_unitario:
            # pylint: disable=E1121
            seletor = Controlador(
                chave_unitario[0],
                corte,
                relig,
                chave_unitario[3],
                num_tse_linhas,
                chave_unitario[4],
                identificador,
                posicao_rede,
                profundidade,
                session
            )
            seletor.executar_processo()
            etapa_unitario.append(chave_unitario[0])

    if chave_rb_investimento:
        session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
        cesta_dicionario.cesta(
            chave_rb_investimento[3],
            chave_rb_investimento[4],
            chave_rb_investimento,
            mae,
            session
        )

    if list_chave_rb_despesa:
        session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABV").select()
        for chave_rb_despesa in list_chave_rb_despesa:
            cesta_dicionario.cesta(
                chave_rb_despesa[3],
                chave_rb_despesa[4],
                chave_rb_despesa,
                mae,
                session
            )

    return (
        tse_proibida,
        list_chave_rb_despesa,
        list_chave_unitario,
        chave_rb_investimento,
        chave_unitario,
        ligacao_errada,
        profundidade_errada
    )
