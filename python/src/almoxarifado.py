# almoxarifado.py
"""Módulo dos materiais contratada e SABESP."""
import pandas as pd
import pywintypes
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


class Almoxarifado:
    """Área de todos materiais obrigatórios por TSE."""

    def __init__(self,
                 hidro,
                 operacao,
                 identificador,
                 diametro_ramal,
                 diametro_rede,
                 contrato,
                 estoque,
                 posicao_rede,
                 session,
                 ) -> None:
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

    def aba_materiais(self):
        """Função habilita aba de materiais no sap"""
        console = Console()
        console.print("Processo de Materiais",
                      style="bold red underline", justify="center")
        self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM").select()
        tb_materiais = self.session.findById(
            "wnd[0]/usr/tabsTAB_ITENS_PRECO/tabpTABM/ssubSUB_TAB:"
            + "ZSBMM_VALORACAO_NAPI:9030/cntlCC_MATERIAIS/shellcont/shell")

        return tb_materiais

    def materiais_vinculados(self, tb_materiais):
        """Retorna um DataFrame com os materiais incluídos na ordem."""
        try:
            sap_material = tb_materiais.GetCellValue(0, "MATERIAL")
            print("Tem material vinculado.")
            num_material_linhas = tb_materiais.RowCount
            lista_data = []
            for i in track(range(num_material_linhas), description="[yellow]Obtendo Materiais..."):
                sap_material = tb_materiais.GetCellValue(
                    i, "MATERIAL")
                sap_etapa_material = tb_materiais.GetCellValue(
                    i, "ETAPA")
                sap_desc_material = tb_materiais.GetCellValue(
                    i, "DESC_MAT")
                sap_qtde_material = tb_materiais.GetCellValue(
                    i, "QUANT")
                data = {"Etapa": sap_etapa_material,
                        "Material": sap_material,
                        "Descrição": sap_desc_material,
                        "Quantidade": sap_qtde_material}
                lista_data.append(data)
            df_materiais = pd.DataFrame(lista_data)
            df_materiais["Quantidade"] = df_materiais["Quantidade"].replace(
                ",", ".", regex=True).astype(float)

            return df_materiais

        # pylint: disable=E1101
        except pywintypes.com_error:
            print("Sem material vinculado.")
            return None

    def inspecao(self, tb_materiais, df_materiais):
        print("Iniciando inspeção de materiais.")
        """Seleciona a Classe da TSE correta."""
        sondagem = [
            "591000",
            "567000",
            "321000",
            "283000",
        ]
        if self.identificador[2] in sondagem:
            materiais_contratada.materiais_contratada(
                tb_materiais, self.contrato,
                self.estoque, self.session)
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
                    print("Aplicando a receita de hidrômetro.")
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
                    print("Aplicando a receita de hidrômetro.")
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
                    print("Aplicando a receita de Cavalete.")
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
                    print("Aplicando a receita de religação.")
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
                    print("Aplicando a receita de Supressão.")

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
                    print("Aplicando a receita "
                          + "de TRA.")
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
                    print("Aplicando a receita de ramal de água")
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
                        print(
                            "Aplicando a receita de hidrômetro em ligação de água nova.")
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
                    print("Aplicando a receita de Rede de Água")
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
                    print("Aplicando a receita de Ramal de Esgoto")
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
                    print("Aplicando a receita de Rede de Esgoto")
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
                    print("Aplicando a receita de PNG Esgoto")
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
                    print("Aplicando a receita de Nivelamento de Poço")
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
                    print("Aplicando a receita de Caixa de Parada")
                    material.receita_caixa_de_parada()

                case _:
                    print("Classe não identificada.")
                    return
        return


def materiais(
    hidro_instalado,
    operacao,
    identificador,
    diametro_ramal,
    diametro_rede,
    contrato,
    estoque,
    posicao_rede,
    session,
):
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
