# base.py
'''Base construtora de classes'''
from abc import abstractmethod
from src.unitarios.interfaces import UnitarioInterface


class BaseUnitario(UnitarioInterface):
    '''Construtor comum'''

    def __init__(self, etapa, corte, relig,
                 reposicao, num_tse_linhas, etapa_reposicao,
                 identificador, posicao_rede, profundidade,
                 session
                 ) -> None:
        self.corte = corte
        self.relig = relig
        self.reposicao = reposicao
        self.num_tse_linhas = num_tse_linhas
        self.etapa_reposicao = etapa_reposicao
        self.posicao_rede = posicao_rede
        self.profundidade = profundidade
        self.session = session
        self.identificador = identificador
        self.etapa = etapa

    @abstractmethod
    def processar(self):
        pass

    @abstractmethod
    def pagar(self):
        pass
