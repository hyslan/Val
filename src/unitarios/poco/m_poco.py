"""Módulo Família Poço Unitário."""
from src.lista_reposicao import dict_reposicao
from src.unitarios.localizador import btn_localizador


class Poco:
    """Classe Unitária de Poço."""
    CODIGOS = {
        'NIV_CX_PARADA': ("456111", ("050401", "050402", "451123")),
        'TROCA_CX_PARADA': ("456112", ("050401", "050402", "451123")),
        'NIVELAMENTO': ("456206", "451208", ("456207", "456207", "451208")),
    }

    def __init__(self, etapa, corte, relig, reposicao, num_tse_linhas,
                 etapa_reposicao, identificador, posicao_rede,
                 profundidade, session, preco):
        self.etapa = etapa
        self.corte = corte
        self.relig = relig
        self.reposicao = reposicao
        self.num_tse_linhas = num_tse_linhas
        self.etapa_reposicao = etapa_reposicao
        self.posicao_rede = posicao_rede
        self.profundidade = profundidade
        self.session = session
        self.identificador = identificador
        self.preco = preco

    def reposicoes(self, cod_reposicao: tuple) -> None:
        """Reposições dos serviços de Poço"""
        rep_com_etapa = [(x, y)
                         for x, y in zip(self.reposicao, self.etapa_reposicao)]

        for pavimento in rep_com_etapa:
            operacao_rep = pavimento[1]
            if operacao_rep == '0':
                operacao_rep = '0010'
            # 0 é tse da reposição;
            # 1 é etapa da tse da reposição;
            if pavimento[0] in dict_reposicao['cimentado']:
                preco_reposicao = cod_reposicao[0]
                txt_reposicao = (
                    f"Pago 1 UN de LRP CIM  - CODIGO: {preco_reposicao}")
            if pavimento[0] in dict_reposicao['especial']:
                preco_reposicao = cod_reposicao[1]
                txt_reposicao = (
                    f"Pago 1 UN de LRP ESP  - CODIGO: {preco_reposicao}")
            if pavimento[0] in dict_reposicao['asfalto_frio']:
                preco_reposicao = cod_reposicao[2]
                txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                 + f" - CODIGO: {preco_reposicao}")

            # 4220 é módulo Investimento.

            btn_localizador(self.preco, self.session, preco_reposicao)
            n_etapa = self.preco.GetCellValue(
                self.preco.CurrentCellRow, "ETAPA")

            if not n_etapa == operacao_rep:
                self.preco.pressToolbarButton("&FIND")
                self.session.findById(
                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                self.session.findById(
                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                self.session.findById("wnd[1]").sendVKey(0)
                self.session.findById("wnd[1]").sendVKey(0)
                self.session.findById("wnd[1]").sendVKey(12)

            self.preco.modifyCell(
                self.preco.CurrentCellRow, "QUANT", "1")
            self.preco.setCurrentCell(
                self.preco.CurrentCellRow, "QUANT")
            self.preco.pressEnter()
            print(txt_reposicao)

    def _repor(self, codigos_reposicao):
        if self.reposicao:
            self.reposicoes(codigos_reposicao)

    def _pagar(self, preco_tse: str) -> None:
        """Pagar serviço"""
        btn_localizador(self.preco, self.session, preco_tse)
        self.preco.modifyCell(
            self.preco.CurrentCellRow, "QUANT", "1")
        self.preco.setCurrentCell(
            self.preco.CurrentCellRow, "QUANT")
        self.preco.pressEnter()
        print(f"Pago 1 UN de {preco_tse}")

    def _processar_operacao(self, tipo_operacao: str):
        codigo = self.CODIGOS.get(tipo_operacao)
        if codigo:
            print(
                f"Iniciando processo de pagar {tipo_operacao.replace('_', ' ')}")
            if tipo_operacao == 'NIVELAMENTO':
                self._pagar(codigo[1])
                return

            self._pagar(codigo[0])
            self._repor(codigo[1])

    def niv_cx_parada(self):
        """Método Nivelamento de Caixa de Parada"""
        self._processar_operacao('NIV_CX_PARADA')

    def troca_de_caixa_de_parada(self):
        """Troca/Descobrimento de Caixa de Parada, Válvula de Rede de Água - Código 456112"""
        self._processar_operacao('TROCA_CX_PARADA')

    def nivelamento(self):
        """Nivelamentos com e sem reposição."""
        self._processar_operacao('NIVELAMENTO')
