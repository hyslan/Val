'''Módulo Família Ligação Água Unitário.'''
from src.unitarios.base import BaseUnitario
from src.unitarios.localizador import btn_localizador
from src.lista_reposicao import dict_reposicao


class LigacaoAgua(BaseUnitario):
    '''Ramo de Ligações (Ramal) de água'''
    MND = ('TA', 'EI', 'TO', 'PO')

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
                preco_reposicao = cod_reposicao
                txt_reposicao = (
                    "Pago 1 UN de LRP CIM  - CODIGO: 456471")
            if pavimento[0] in dict_reposicao['especial']:
                preco_reposicao = cod_reposicao
                txt_reposicao = (
                    "Pago 1 UN de LRP ESP  - CODIGO: 456472")
            if pavimento[0] in dict_reposicao['asfalto_frio']:
                preco_reposicao = cod_reposicao
                txt_reposicao = ("Pago 1 UN de LPB ASF MND LAG AVUL COMPX C"
                                 + " - CODIGO: 451495")

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
            contador_pg += 1

    def posicao_pagar(self, preco_tse: str) -> None:
        '''Paga de acordo com a posição da rede'''
        preco = self.preco()
        if preco is not None:
            ramal = False
            if ramal is False:
                # Botão localizar
                btn_localizador(preco, self.session, preco_tse)
                preco.modifyCell(
                    preco.CurrentCellRow, "QUANT", "1")
                preco.setCurrentCell(
                    preco.CurrentCellRow, "QUANT")
                preco.pressEnter()
                print(f"Pago 1 UN de {preco_tse}")
                ramal = True

    def ligacao_agua(self):
        '''Ramal novo de água, avulsa.'''
        if self.posicao_rede == 'PA':
            # LAG sem fornecimento Código: 456451
            # LAG com fornecimento Código: 456461
            print(
                f"Iniciando processo de pagar ligação de água posicao: {self.posicao_rede}"
            )
            self.posicao_pagar("456451")
            if self.reposicao:
                # Ordem da tupla: Cimentado, Especial e Asfalto Frio
                self.reposicoes(("456471", "456472", "451495"))

        if self.posicao_rede in LigacaoAgua.MND:
            # LAG sem fornecimento Código: 456491
            # LAG com fornecimento Código: 456492
            print(
                f"Iniciando processo de pagar ligação de água posicao: {self.posicao_rede}"
            )
            self.posicao_pagar("456491")
            if self.reposicao:
                # Ordem da tupla: Cimentado, Especial e Asfalto Frio
                self.reposicoes(("456493", "456494", "451495"))

    def tra_nv(self):
        '''Troca de Ramal de água não visível'''
        pass

    def png(self):
        '''Passado novo ramal para nova rede - Obra'''
        pass

    def subst_agua(self):
        '''Substituição de ramal de água, tem adicional de suprimir
        o ferrule da rede'''
        pass

    def tra_prev(self):
        '''Troca de ramal de água preventiva'''
        pass
