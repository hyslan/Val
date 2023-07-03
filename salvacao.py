# salvacao.py
'''Módulo de salvar valoração.'''
import pywintypes
from sap_connection import connect_to_sap
from excel_tbs import load_worksheets
from confere_os import consulta_os

session = connect_to_sap()
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


def salvar(ordem, int_num_lordem, qtd_ordem):
    '''Função para salvar valoração.'''
    try:
        session.findById("wnd[0]").sendVKey(11)
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    # pylint: disable=E1101
    except pywintypes.com_error:
        return qtd_ordem
    print("Salvando valoração!")
    # Verificar se Salvou
    (status_sistema,
        status_usuario,
        *_) = consulta_os(ordem)
    print("Verificando se Ordem foi valorada.")
    if status_usuario == "EXEC VALO":
        print(f"Status da Ordem: {status_sistema}, {status_usuario}")
        print("Foi Salvo com sucesso!")
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        selecao_carimbo.value = "VALORADA"
        lista.save('lista.xlsx')  # salva Planilha
        qtd_ordem += 1
    else:
        print(f"Ordem: {ordem} não foi salva.")
        selecao_carimbo = planilha.cell(row=int_num_lordem, column=2)
        selecao_carimbo.value = "NÃO FOI SALVO"
        lista.save('lista.xlsx')  # salva Planilha
    return qtd_ordem
