# controlador.py
"""Classe controladora."""

from python.src.unitarios.base import BaseUnitario
from python.src.unitarios.dicionario import selecionar_tse


class Controlador(BaseUnitario):
    """Seletor de classes de família de serviços."""

    def preserv_inter_serv(self) -> None:
        """Etapa de preservação de Interferência."""
        ramo_agua = [
            ["hidrometro"],
            ["cavalete"],
            ["religacao"],
            ["supressao"],
            ["ramal_agua"],
            ["tra"],
            ["ligacao_agua"],
            ["ligacao_agua_nova"],
            ["rede_agua"],
            ["supr_restab"],
        ]
        self.identificador.remove("preservacao")
        serv_preservacao = "456118" if self.identificador in ramo_agua else "456217"

        preco = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABI/ssubSUB_TAB:ZSBMM_VALORACAO_NAPI:9020/cntlCC_ITEM_PRECO/shellcont/shell",
        )
        preco.GetCellValue(0, "NUMERO_EXT")
        if preco is not None:
            num_precos_linhas = preco.RowCount
            n_preco = 0  # índice para itens de preço
            for n_preco, sap_preco in enumerate(range(num_precos_linhas)):
                sap_preco = preco.GetCellValue(n_preco, "NUMERO_EXT")
                if sap_preco == serv_preservacao:
                    # Marca pagar na TSE
                    preco.modifyCell(n_preco, "QUANT", "1")
                    preco.setCurrentCell(n_preco, "QUANT")
                    preco.pressEnter()
                    break

    def _processar_operacao(self, tipo_operacao: str) -> None:
        """Processar Código de preço."""

    def executar_processo(self) -> None:
        """Selecionar a classe apropriada com base no código da etapa."""
        if self.etapa in ("713000", "713500"):
            self.preserv_inter_serv()

        classe_unitario = selecionar_tse(
            self.etapa,
            self.corte,
            self.relig,
            self.reposicao,
            self.num_tse_linhas,
            self.etapa_reposicao,
            self.identificador,
            self.posicao_rede,
            self.profundidade,
            self.session,
            self.preco,
        )

        if classe_unitario:
            # Processar e pagar
            classe_unitario()
        else:
            pass
