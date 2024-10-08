# almoxarifado.py
"""Módulo dos materiais contratada e SABESP."""

import pandas as pd
import pywintypes
import win32com.client as win32
from rich.console import Console
from rich.progress import track

from python.src.wms import (
    cavalete_material,
    corte_restab_material,
    hidrometro_material,
    materiais_contratada,
    poco_material,
    rede_agua_material,
    rede_esgoto_material,
)

from .log_decorator import log_execution


@log_execution
class Almoxarifado:
    """Área de todos materiais obrigatórios por TSE."""

    def __init__(
        self,
        hidro: str,
        operacao: str,
        identificador: tuple[str, str, str, list[str], list[str]],
        diametro_ramal: str,
        diametro_rede: str,
        contrato: str,
        estoque: pd.DataFrame,
        posicao_rede: str,
        session: win32.CDispatch,
    ) -> None:
        """Inicializa a classe Almoxarifado.

        Args:
        ----
            hidro (str): Número de Série do Hidro
            operacao (str): Etapa Pai
            identificador (tuple[str, str, str]): TSE, Etapa TSE, ID Match Case
            diametro_ramal (str): Tamanho do ramal
            diametro_rede (str): Tamanho da rede
            contrato (str): Número do Contrato
            estoque (pd.DataFrame): Df com o estoque do contrato
            posicao_rede (str): Posição de PA até PO
            session (win32.CDispatch): Sessão do SAPGUI

        """
        self.hidro = hidro
        # 0 - tse, 1 - etapa tse, 2 - id match case
        self.identificador = identificador
        self.operacao = operacao
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.contrato = contrato
        self.estoque = estoque
        self.posicao_rede = posicao_rede
        self.session = session

    def aba_materiais(self) -> win32.CDispatch:
        """Função habilita aba de materiais no sap."""
        console = Console()
        console.print("Processo de Materiais", style="bold red underline", justify="center")
        self.session.findById("wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM").select()
        return self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM/ssubSUB_TAB:ZSBMM_VALORACAO_NAPI:9030/cntlCC_MATERIAIS/shellcont/shell",
        )

    def materiais_vinculados(self, tb_materiais: win32.CDispatch) -> pd.DataFrame:
        """Retorna um DataFrame com os materiais incluídos na ordem."""
        try:
            sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
            num_material_linhas = tb_materiais.RowCount
            lista_data = []
            for i in track(range(num_material_linhas), description="[yellow]Obtendo Materiais..."):
                sap_material = tb_materiais.GetCellValue(i, "MATERIAL")
                sap_etapa_material = tb_materiais.GetCellValue(i, "ETAPA")
                sap_desc_material = tb_materiais.GetCellValue(i, "DESC_MAT")
                sap_qtde_material = tb_materiais.GetCellValue(i, "QUANT")
                data = {
                    "Etapa": sap_etapa_material,
                    "Material": sap_material,
                    "Descrição": sap_desc_material,
                    "Quantidade": sap_qtde_material,
                }
                lista_data.append(data)
            df_materiais = pd.DataFrame(lista_data)
            df_materiais["Quantidade"] = df_materiais["Quantidade"].replace(",", ".", regex=True).astype(float)

        # pylint: disable=E1101
        except pywintypes.com_error:
            return pd.DataFrame(columns=["Etapa", "Material", "Descrição", "Quantidade"]).astype(
                {"Etapa": object, "Material": object, "Descrição": object, "Quantidade": float},
            )
        else:
            return df_materiais

    # ruff: noqa: C901
    def inspecao(self, tb_materiais: win32.CDispatch, df_materiais: pd.DataFrame) -> None:
        """Seleciona a Classe da TSE correta."""
        sondagem = [
            "591000",
            "567000",
            "321000",
            "283000",
        ]
        if self.identificador[2] in sondagem:
            materiais_contratada.materiais_contratada(tb_materiais, self.contrato, self.estoque, self.session)
        else:
            match self.identificador[2]:
                case "hidrometro":
                    material = hidrometro_material.HidrometroMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )
                    material.receita_hidrometro()

                case "desinclinado":
                    material = hidrometro_material.HidrometroMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )
                    material.receita_desinclinado_hidrometro()

                case "cavalete":
                    material = cavalete_material.CavaleteMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )
                    material.receita_cavalete()

                case "religacao":
                    material = corte_restab_material.CorteRestabMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )
                    material.receita_religacao()

                case "supressao":
                    material = corte_restab_material.CorteRestabMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )

                    material.receita_supressao()
                case "ramal_agua" | "tra":
                    material = rede_agua_material.RedeAguaMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        df_materiais,
                        self.posicao_rede,
                        self.session,
                    )
                    material.receita_tra()

                case "reparo_ramal_agua" | "ligacao_agua" | "ligacao_agua_nova":
                    material = rede_agua_material.RedeAguaMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        df_materiais,
                        self.posicao_rede,
                        self.session,
                    )
                    material.receita_reparo_de_ramal_de_agua()
                    # Olhar hidrômetro e lacre em ligações novas.
                    if self.identificador[2] == "ligacao_agua_nova":
                        material = hidrometro_material.HidrometroMaterial(
                            self.hidro,
                            self.operacao,
                            self.identificador,
                            self.diametro_ramal,
                            self.diametro_rede,
                            tb_materiais,
                            self.contrato,
                            self.estoque,
                            self.session,
                        )
                        material.receita_hidrometro()

                case "rede_agua" | "gaxeta" | "chumbo_junta" | "valvula":
                    material = rede_agua_material.RedeAguaMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        df_materiais,
                        self.posicao_rede,
                        self.session,
                    )
                    material.receita_reparo_de_rede_de_agua()

                case "ligacao_esgoto":
                    material = rede_esgoto_material.RedeEsgotoMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        df_materiais,
                        self.posicao_rede,
                        self.session,
                    )
                    material.receita_reparo_de_ramal_de_esgoto()

                case "rede_esgoto":
                    material = rede_esgoto_material.RedeEsgotoMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        df_materiais,
                        self.posicao_rede,
                        self.session,
                    )
                    material.receita_reparo_de_rede_de_esgoto()

                case "png_esgoto":
                    material = rede_esgoto_material.RedeEsgotoMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        df_materiais,
                        self.posicao_rede,
                        self.session,
                    )
                    material.png()

                case "preservacao":
                    pass

                case "poço":
                    material = poco_material.PocoMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )
                    material.niv_pv_pi()

                case "cx_parada":
                    material = poco_material.PocoMaterial(
                        self.hidro,
                        self.operacao,
                        self.identificador,
                        self.diametro_ramal,
                        self.diametro_rede,
                        tb_materiais,
                        self.contrato,
                        self.estoque,
                        self.session,
                    )
                    material.receita_caixa_de_parada()

                case _:
                    return
        return


def materiais(
    hidro_instalado: str,
    operacao: str,
    identificador: tuple[str, str, str, list[str], list[str]],
    diametro_ramal: str,
    diametro_rede: str,
    contrato: str,
    estoque: pd.DataFrame,
    posicao_rede: str,
    session: win32.CDispatch,
) -> None:
    """Função dos materiais de acordo com a TSE pai."""
    servico = Almoxarifado(
        hidro_instalado,
        operacao,
        identificador,
        diametro_ramal,
        diametro_rede,
        contrato,
        estoque,
        posicao_rede,
        session,
    )
    tb_materiais = servico.aba_materiais()
    df_materiais = servico.materiais_vinculados(tb_materiais)

    servico.inspecao(tb_materiais, df_materiais)
