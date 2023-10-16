'''Módulo de busca código de preço'''


def btn_localizador(preco, session, codigo):
    '''Objeto localizar.'''
    preco.pressToolbarButton("&FIND")
    session.findById(
        "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = codigo
    session.findById(
        "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]").sendVKey(12)
