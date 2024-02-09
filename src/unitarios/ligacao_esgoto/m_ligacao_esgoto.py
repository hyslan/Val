'''Módulo Família Ligação Água Unitário.'''

from src.unitarios.localizador import btn_localizador
from src.lista_reposicao import dict_reposicao


class LigacaoEsgoto:
    '''Ramo de Ligações (Ramal) de água'''

    # Ordem da tupla: [0] -> preço s/ fornecimento, [1] -> c/ fornecimento,
    # [2] Reposições -> (Cimentado, Especial e Asfalto Frio)
    # Códigos de preço para posição de rede do serviço pai
    # Ver profundidade até 2m, 2m a 3m e 3m a 4m
    CODIGOS_2M = {
        'LESG_PA': ("456651", "456671", ("456711", "456712", "451713")),
        'LESG_TA': ("456652", "456672", ("456711", "456712", "451713")),
        'LESG_EI': ("456653", "456673", ("456711", "456712", "451716")),
        'LESG_TO': ("456654", "456674", ("456711", "456712", "451719")),
        'LESG_PO': ("456655", "456675", ("456711", "456712", "451719")),
        'TRE_PA': ("457001", "457021", ("457101", "457104", "452107")),
        'TRE_TA': ("457002", "457022", ("457101", "457104", "452107")),
        'TRE_EI': ("457003", "457023", ("457101", "457104", "452116")),
        'TRE_TO': ("457004", "457024", ("457101", "457104", "452125")),
        'TRE_PO': ("457005", "457025", ("457101", "457104", "452125"))
    }
    CODIGOS_3M = {
        'LESG_PA': ("456656", "456676", ("456722", "456723", "451724")),
        'LESG_TA': ("456657", "456677", ("456722", "456723", "451724")),
        'LESG_EI': ("456658", "456678", ("456722", "456723", "451727")),
        'LESG_TO': ("456659", "456679", ("456722", "456723", "451730")),
        'LESG_PO': ("456660", "456680", ("456722", "456723", "451730")),
        'TRE_PA': ("457006", "457026", ("457102", "457105", "452108")),
        'TRE_TA': ("457007", "457027", ("457102", "457105", "452108")),
        'TRE_EI': ("457008", "457028", ("457102", "457105", "452117")),
        'TRE_TO': ("457009", "457029", ("457102", "457105", "452126")),
        'TRE_PO': ("457010", "457030", ("457102", "457105", "452126"))
    }
    CODIGOS_4M = {
        'TRE_PA': ("457011", "457031", ("457103", "457105", "452113")),
        'TRE_TA': ("457012", "457032", ("457103", "457105", "452113")),
        'TRE_EI': ("457013", "457033", ("457103", "457105", "452118")),
        'TRE_TO': ("457014", "457034", ("457103", "457105", "452127")),
        'TRE_PO': ("457015", "457035", ("457103", "457105", "452127"))
    }
    CODIGOS_PNG = {
        'PNG_PA': ("456371", "456381", ("456631", "456632", "451633")),
        'PNG_TA': ("456411", "456881", ("456631", "456632", "451633")),
        'PNG_EI': ("456411", "456881", ("456631", "456632", "451636")),
        'PNG_TO': ("456411", "456881", ("456631", "456632", "451639")),
        'PNG_PO': ("456411", "456881", ("456631", "456632", "451639"))
    }

    def __init__(self, etapa, corte, relig, reposicao, num_tse_linhas,
                 etapa_reposicao, identificador, posicao_rede, profundidade, session):
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
        self._ramal = False

    def preco(self):
        '''Shell do itens preço'''
        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:"
            + "ZSBMM_VALORACAOINV:9020/cntlCC_ITEM_PRECO/shellcont/shell")
        preco.GetCellValue(0, "NUMERO_EXT")
        return preco

    def reposicoes(self, cod_reposicao: tuple) -> None:
        '''Reposições dos serviços de Ligação de água'''
        preco = self.preco()
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

            btn_localizador(preco, self.session, preco_reposicao)
            n_etapa = preco.GetCellValue(
                preco.CurrentCellRow, "ETAPA")

            if not n_etapa == operacao_rep:
                preco.pressToolbarButton("&FIND")
                self.session.findById(
                    "wnd[1]/usr/txtGS_SEARCH-VALUE").Text = preco_reposicao
                self.session.findById(
                    "wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
                self.session.findById("wnd[1]").sendVKey(0)
                self.session.findById("wnd[1]").sendVKey(0)
                self.session.findById("wnd[1]").sendVKey(12)

            preco.modifyCell(
                preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(
                preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print(txt_reposicao)

    def _posicao_pagar(self, preco_tse: str) -> None:
        '''Paga de acordo com a posição da rede'''
        if not self._ramal:
            preco = self.preco()
            btn_localizador(preco, self.session, preco_tse)
            preco.modifyCell(
                preco.CurrentCellRow, "QUANT", "1")
            preco.setCurrentCell(
                preco.CurrentCellRow, "QUANT")
            preco.pressEnter()
            print(f"Pago 1 UN de {preco_tse}")
            self._ramal = True

    def _repor(self, codigos_reposicao):
        if self.reposicao:
            self.reposicoes(codigos_reposicao)

    def _processar_operacao(self, tipo_operacao: str) -> None:
        try:
            profundidade_float = float(self.profundidade.replace(',', '.'))
            if tipo_operacao == "PNG":
                codigos = self.CODIGOS_PNG
            if profundidade_float <= 2.00:
                codigos = self.CODIGOS_2M
            if profundidade_float > 2.00 and profundidade_float <= 3.00:
                codigos = self.CODIGOS_3M
            if tipo_operacao == "TRE" and profundidade_float > 3.00:
                codigos = self.CODIGOS_4M
        except ValueError:
            return

        match self.posicao_rede:
            case "PA":
                codigo = codigos.get(tipo_operacao + '_PA')
            case "TA":
                codigo = codigos.get(tipo_operacao + '_TA')
            case "EI":
                codigo = codigos.get(tipo_operacao + '_EI')
            case "TO":
                codigo = codigos.get(tipo_operacao + '_TO')
            case "PO":
                codigo = codigos.get(tipo_operacao + '_PO')
            case _:
                return

        if codigo:
            print(
                f"Iniciando processo de pagar {tipo_operacao.replace('_', ' ')}"
                " posição: {self.posicao_rede}")
            self._posicao_pagar(codigo[0])
            self._repor(codigo[2])

    def ligacao_esgoto(self):
        '''Ramal novo de água, avulsa.'''
        if self.posicao_rede:
            self._processar_operacao('LESG')

    def tre(self):
        '''Troca de Ramal de água não visível'''
        if self.posicao_rede:
            self._processar_operacao('TRE')

    def png(self):
        '''Passado novo ramal para nova rede - Obra'''
        if self.posicao_rede:
            self._processar_operacao('PNG')
