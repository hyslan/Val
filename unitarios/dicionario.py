import sys


def Unitario(Etapa, Corte, Relig):
    from unitarios.hidrometro import m_hidrometro
    from unitarios.supressao import m_supressao
    from unitarios.religacao import m_religacao
    from unitarios.cavalete import m_cavalete
    DicionarioUN = {
        '201000': m_hidrometro.Hidrometro.THD_456901,
        '203000': m_hidrometro.Hidrometro.THD_456901,
        '203500': m_hidrometro.Hidrometro.HD_456022,
        '204000': m_hidrometro.Hidrometro.THD_456901,
        '205000': m_hidrometro.Hidrometro.THD_456901,
        '206000': m_hidrometro.Hidrometro.THD_456901,
        '207000': m_hidrometro.Hidrometro.THD_456901,
        '215000': m_hidrometro.Hidrometro.THDPrev_456902,
        '405000': m_supressao.Corte.Supressao,
        '414000': m_supressao.Corte.Supressao,
        '450500': m_religacao.Religacao.Restabelecida,
        '453000': m_religacao.Religacao.Restabelecida,
        '455500': m_religacao.Religacao.Restabelecida,
        '463000': m_religacao.Religacao.Restabelecida,
        '465000': m_religacao.Religacao.Restabelecida,
        '467500': m_religacao.Religacao.Restabelecida,
        '475500': m_religacao.Religacao.Restabelecida,
        # '142000': m_cavalete.Cavalete.
        # '148000': m_cavalete.Cavalete.
        '149000': m_cavalete.Cavalete.TrocaCvKit,
        '153000': m_cavalete.Cavalete.TrocaPeCvPrev,

        # Adicionar chaves conforme classes e métodos forem adicionados ao diretório e instanciados na main
    }

    if Etapa in DicionarioUN:
        print(f"Etapa está inclusa no Dicionário de Unitários: {Etapa}")
        metodo = DicionarioUN[Etapa]
        metodo(Corte, Relig)  # Chama o método de uma classe dentro do Dicionário
    else:
        print("TSE não Encontrada no Dicionário!")
        sys.exit()
