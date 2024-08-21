# base.py
"""Base construtora de classes"""
from python.src.unitarios.interfaces import UnitarioInterface


class BaseUnitario(UnitarioInterface):
    """Construtor comum"""

    def __init__(self, etapa, corte, relig,
                 reposicao, num_tse_linhas, etapa_reposicao,
                 identificador, posicao_rede, profundidade,
                 session,
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
        self.preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAO_NAPI:9020/cntlCC_ITEM_PRECO/shellcont/shell")

    def processar(self):
        pass

    def pagar(self):
        pass

    def _processar_operacao(self, tipo_operacao):
        """Processar Código de preço"""
