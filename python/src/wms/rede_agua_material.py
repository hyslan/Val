# hidrometro_material.py
"""Módulo dos materiais de família Rede de Água."""

from __future__ import annotations

import typing

from rich.console import Console

from python.src.wms import localiza_material, materiais_contratada, testa_material_sap

if typing.TYPE_CHECKING:
    from pandas import DataFrame
    from win32com.client import CDispatch

console: Console = Console()


class RedeAguaMaterial:
    """Classe de materiais de CRA."""

    def __init__(
        self,
        hidro: str,
        operacao: str,
        identificador: tuple[str, str, str],  # Unique Array key
        diametro_ramal: str,
        diametro_rede: str,
        tb_materiais: CDispatch,
        contrato: str,
        estoque: DataFrame,
        session: CDispatch,
    ) -> None:
        """Método de inicialização da classe Rede Água.

        Args:
        ----
            hidro (str): Número de Série do Hidrometro.
            operacao (str): Etapa Pai
            identificador (tuple[str, str, str]): TSE, Etapa TSE, ID Match Case do inspector de Almoxarixado.py
            diametro_ramal (str): Tamanho do diâmetro do ramal.
            diametro_rede (str): Tamanho do diâmetro da rede.
            tb_materiais (CDispatch): GRID de Materiais.
            contrato (str): Número do contrato.
            estoque (DataFrame): Tabela de Estoque.
            session (CDispatch): Sessão do SAP.

        """
        self.hidro = hidro
        self.operacao = operacao
        self.identificador = identificador
        self.diametro_ramal = diametro_ramal
        self.diametro_rede = diametro_rede
        self.tb_materiais = tb_materiais
        self.contrato = contrato
        self.estoque = estoque
        self.df_materiais = df_materiais
        self.posicao_rede = posicao_rede
        self.session = session
        self.list_contratada = materiais_contratada.lista_materiais()
        self.usr = session.findById("wnd[0]/usr")

    def materiais_vigentes(self) -> None:
        """Materiais com estoque."""
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)
        # CONEXOES MET LIGACOES  FÊMEA DN 20
        con_met_femea_estoque = self.estoque[self.estoque["Material"] == "30002394"]
        # CONEXOES MET ADAP MACHO DN 20
        self.estoque[self.estoque["Material"] == "30001346"]
        # DISPOSITIVO MED PLASTICO DN 20
        self.estoque[self.estoque["Material"] == "50000178"]
        # REGISTRO METALICO RAMAL PREDIAL DN 20
        self.estoque[self.estoque["Material"] == "30006747"]
        # COLAR TOM P/TUBO PE DE 32XDN 20 TE INTEG
        self.estoque[self.estoque["Material"] == "30000287"]
        # COLAR TOMADA ACO INOX DN50A150 X DNR20
        colar_tom_aco_inox__dn50a150xdnr20_estoque = self.estoque[self.estoque["Material"] == "30004702"]
        # COLAR TOMADA ACO INOX DN200A300 X DNR20
        self.estoque[self.estoque["Material"] == "30004701"]
        # COLAR TOMADA FF C.INOX DN200A300 X DNR20
        self.estoque[self.estoque["Material"] == "30004703"]
        # ABRACADEIRA FF REPARO TUBO DN75 LMIN=150
        abrac_ff_reparo_dn75_lmin150_estoque = self.estoque[self.estoque["Material"] == "30001122"]
        # ABRACADEIRA INOX REPARO TUBO DN50 L=300
        abrac_inox_reparo_dn50_l300_estoque = self.estoque[self.estoque["Material"] == "30002151"]
        # TUBO PEAD DN 20
        self.estoque[self.estoque["Material"] == "30001848"]

        if sap_material is not None:
            num_material_linhas = self.tb_materiais.RowCount
            ultima_linha_material = num_material_linhas
            self.df_materiais[self.df_materiais["Material"] == "30001848"]
            self.df_materiais[self.df_materiais["Material"] == "30002394"]
            self.df_materiais[self.df_materiais["Material"] == "30001346"]
            self.df_materiais[self.df_materiais["Material"] == "50000178"]
            self.df_materiais[self.df_materiais["Material"] == "30006747"]
            self.df_materiais[self.df_materiais["Material"] == "30000287"]
            self.df_materiais[self.df_materiais["Material"] == "30004702"]
            self.df_materiais[self.df_materiais["Material"] == "30004701"]
            self.df_materiais[self.df_materiais["Material"] == "30004703"]

            for material in self.df_materiais["Material"]:
                match material:
                    case "30002394":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30002394")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "30001346":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30001346")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "50000178":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "50000178")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "30006747":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30006747")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "30000287":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30000287")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "30004702":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30004702")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "30004701":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30004701")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    case "30004703":
                        resultado = localiza_material.qtd_max(material, self.estoque, 1, self.df_materiais)
                        if not resultado.empty:
                            localiza_material.btn_busca_material(self.tb_materiais, self.session, "30004703")
                            localiza_material.qtd_correta(self.tb_materiais, "1")
                    # TUBO PEAD DN20
                    case "30001848":
                        codigo = "30001848"
                        match self.posicao_rede:
                            case "PA":
                                resultado = localiza_material.qtd_max(material, self.estoque, 2.5, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(self.tb_materiais, self.session, codigo)
                                    localiza_material.qtd_correta(self.tb_materiais, "2")
                            case "TA":
                                resultado = localiza_material.qtd_max(material, self.estoque, 4.5, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(self.tb_materiais, self.session, codigo)
                                    localiza_material.qtd_correta(self.tb_materiais, "4")
                            case "EX":
                                resultado = localiza_material.qtd_max(material, self.estoque, 10, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(self.tb_materiais, self.session, codigo)
                                    localiza_material.qtd_correta(self.tb_materiais, "5")
                            case "TO":
                                resultado = localiza_material.qtd_max(material, self.estoque, 15, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(self.tb_materiais, self.session, codigo)
                                    localiza_material.qtd_correta(self.tb_materiais, "7")
                            case "PO":
                                resultado = localiza_material.qtd_max(material, self.estoque, 15, self.df_materiais)
                                if not resultado.empty:
                                    localiza_material.btn_busca_material(self.tb_materiais, self.session, codigo)
                                    localiza_material.qtd_correta(self.tb_materiais, "10")

            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(n_material, "ETAPA")

                match sap_material:
                    case "30029526":
                        if self.contrato == "4600041302":
                            # UNIÃO PLASTICA P/ TUBO PE DE 20
                            pass

                    case "30001122":
                        if abrac_ff_reparo_dn75_lmin150_estoque.empty:
                            self.tb_materiais.modifyCheckbox(
                                n_material,
                                "ELIMINADO",
                                True,
                            )

                    case "30002735":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30011136":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30005088":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )
                        if not con_met_femea_estoque.empty:
                            self.tb_materiais.InsertRows(str(ultima_linha_material))
                            self.tb_materiais.modifyCell(
                                ultima_linha_material,
                                "ETAPA",
                                sap_etapa_material,
                            )
                            # Adiciona CONEXÕES METALICAS COTOVELO FEMEA DN 3/4.
                            self.tb_materiais.modifyCell(
                                ultima_linha_material,
                                "MATERIAL",
                                "30002394",
                            )
                            self.tb_materiais.modifyCell(
                                ultima_linha_material,
                                "QUANT",
                                "1",
                            )
                            self.tb_materiais.setCurrentCell(
                                ultima_linha_material,
                                "QUANT",
                            )
                            ultima_linha_material = ultima_linha_material + 1

                    case "30004097":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30002152":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30002151":
                        if abrac_inox_reparo_dn50_l300_estoque.empty:
                            self.tb_materiais.modifyCheckbox(
                                n_material,
                                "ELIMINADO",
                                True,
                            )

                    case "30002802":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30004104":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30007931":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )

                    case "30002202":
                        self.tb_materiais.modifyCheckbox(
                            n_material,
                            "ELIMINADO",
                            True,
                        )
                        if not colar_tom_aco_inox__dn50a150xdnr20_estoque.empty:
                            self.tb_materiais.InsertRows(str(ultima_linha_material))
                            self.tb_materiais.modifyCell(
                                ultima_linha_material,
                                "ETAPA",
                                sap_etapa_material,
                            )
                            # Adiciona COLAR TOMADA ACO INOX DN50A150 X DNR20.
                            self.tb_materiais.modifyCell(
                                ultima_linha_material,
                                "MATERIAL",
                                "30004702",
                            )
                            self.tb_materiais.modifyCell(
                                ultima_linha_material,
                                "QUANT",
                                "1",
                            )
                            self.tb_materiais.setCurrentCell(
                                ultima_linha_material,
                                "QUANT",
                            )
                            ultima_linha_material = ultima_linha_material + 1

    def receita_reparo_de_rede_de_agua(self) -> None:
        """Padrão de materiais na classe CRA."""
        materiais_receita = [
            "10014709",
            # ABRACADEIRA FF -> TRIPARTIDA
            "30008103",
            "30002141",
            "30001120",
            "30004084",
            "30002142",
            "30000255",
            "30004091",
            "30001122",
            # CAP ELETROFUSÃO
            "30002352",
            "30004405",
            # CAP FF DUCTIL BOLSA
            "30001055",
            # CAP FF PARA PVC
            "30002356",
            "30000564",
            # CAP PVC
            "30004421",
            # COLAR TOMADA
            "30000287",
            "30004701",
            "30004702",
            "30002202",
            # COLARINHO
            "30004732",
            "30000249",
            "30004556",
            # CURVA 45 GR FF P/ PVC
            "30005264",
            "30005265",
            "30005266",
            # CURVA 45 GR PVC
            "30005280",
            "30005283",
            "30005282",
            # CURVA 90 GR FF P/ PVC
            "30002714",
            "30005330",
            # CURVA 90GR FF BOLSAS
            "30008324",
            # EXTREMIDADE FF P FLG
            "30005517",
            "30005519",
            # JUNTA FF BOL VAR DIA
            "30008409",
            "30005611",
            # LUVA BIPARTIDA FF
            "30005887",
            "30001546",
            "30005888",
            "30002794",
            "30005889",
            "30002795",
            "30005889",
            "30005893",
            # LUVA CORRER FF C/ BOLSAS
            "30005906",
            "30002803",
            "30005917",
            "30002809",
            "30005919",
            "30005920",
            "30008431",
            "30005924",
            "30002812",
            # LUVA CORRER PVC BB
            "30008433",
            "30001557",
            # LUVA ELETROFUSÃO
            "30002823",
            "30005752",
            "30002826",
            # LUVA FF PARA PVC
            "30005772",
            "30005773",
            # REDUÇÃO CONC ELETROFUSÃO
            "30001590",
            # REDUÇÃO FF PONTA BOLSA
            "30006911",
            # TE ELETROFUSÃO
            "30008671",
            # TE DE SELA ELETROFUSÃO
            "30011078",
            # TE DE SERVIÇO INTERGRADO ARTICULADO
            "30000211",
            "30007034",
            "30007235",
            # TE FF DUCTIL C/ BOLSAS
            "30008695",
            "30007286",
            "30003716",
            # TE FF PARA PVC
            "30001914",
            "30003730",
            "30000213",
            # TE REDUÇÃO ELETROFUSÃO
            "30007228",
            # TUBO
            "30007853",
            "30028866",
            "30028862",
            "30028864",
            "30007896",
            # UNIÃO P/ PEAD
            "30003832",
            "30007765",
            # UNIÃO PLASTICA P/ TUBO PE DE 20
            "30029526",
        ]
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)

        if sap_material is None:
            return
        material_lista: list[dict[str, str]] = []
        num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
        # Loop do Grid Materiais.
        for n_material in range(num_material_linhas):
            # Pega valor da célula 0
            sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
            sap_etapa_material = self.tb_materiais.GetCellValue(n_material, "ETAPA")
            material_lista.append({"Material": sap_material, "Etapa": sap_etapa_material})
            material_estoque = self.estoque[self.estoque["Material"] == sap_material]

            console.print(f"\n{material_estoque}", style="italic green")

            if (
                sap_material not in materiais_receita
                and sap_material not in self.list_contratada
                and sap_etapa_material == self.operacao
            ):
                self.tb_materiais.modifyCheckbox(
                    n_material,
                    "ELIMINADO",
                    True,
                )
            if material_estoque.empty:
                self.tb_materiais.modifyCheckbox(
                    n_material,
                    "ELIMINADO",
                    True,
                )

        self.materiais_vigentes()

        # Materiais do Global.
        materiais_contratada.materiais_contratada(self.tb_materiais, self.contrato, self.estoque, self.session)

    def receita_tra(self) -> None:
        """Padrão de materiais na classe Troca de Conexão de Ligação de Água."""
        materiais_receita = [
            "30007896",
            "50001070",
            "30001348",
            "30006747",
            "50000021",
            "10014709",
            # --- TE DE SERVIÇO INTERGRADO DN 20 OU 32 ---
            "30008677",  # ART DN 100 P/ DE 32
            "30000211",  # ART DN 100-DE 110 X 20
            "30007034",  # ART DN 50 P/ DE 20
            "30003467",  # ART DN 50 P/ D3 32
            "30007235",  # ART DN 75 P/ DE 20
            "30001683",  # ART DN 75 P/ DE 32
            # --- TE DE SERVIÇO ELETROFUSÃO DN 20 OU 32 ---
            "30000992",  # DE 110 X 20
            "30003669",  # DE 160 X 20
            "30007240",  # DE 225 X 20
            "30007245",  # DE 90 X 32
            "30003672",  # DE 90 X 20
            "30007195",  # TE PP P/ TUBO PEAD DE 32 X 32 MM
            # --- COLAR TOMADA ACO INOX DNR20 ---
            "30004701",  # DN 200 A 300
            "30004702",  # DN 50 A 150
            "30004703",  # DN 200 A 300
            # --- COLAR TOMADA FERRO CINTA INOX ---
            "30002202",  # DN 50 A 150 X DNR25
            "30002204",  # DN 100 X DNR50
            "30001080",  # DN 400 X DNR20
        ]
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)
        # CONEXOES MET LIGACOES FEMEA DN 20
        con_met_femea_dn20_estoque = self.estoque[self.estoque["Material"] == "30002394"]
        # REGISTRO METALICO RAMAL PREDIAL DN 20
        reg_met_predial_dn20_estoque = self.estoque[self.estoque["Material"] == "30006747"]
        if sap_material is None:
            ultima_linha_material = 0
            if not con_met_femea_dn20_estoque.empty:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "ETAPA",
                    self.identificador[1],
                )
                # Adiciona CONEXOES MET LIGACOES FEMEA DN 20.
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "MATERIAL",
                    "30002394",
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "QUANT",
                    "1",
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material,
                    "QUANT",
                )
                ultima_linha_material = ultima_linha_material + 1

            if not reg_met_predial_dn20_estoque.empty:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "ETAPA",
                    self.identificador[1],
                )
                # Adiciona REGISTRO METALICO RAMAL PREDIAL DN 20.
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "MATERIAL",
                    "30006747",
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "QUANT",
                    "1",
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material,
                    "QUANT",
                )
                ultima_linha_material = ultima_linha_material + 1
        else:
            material_lista: list[dict[str, str]] = []
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            # Número da Row do Grid Materiais do SAP
            n_material = 0
            ultima_linha_material = num_material_linhas

            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                # Pega valor da célula 0
                sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(n_material, "ETAPA")
                material_lista.append({"Material": sap_material, "Etapa": sap_etapa_material})
                material_estoque = self.estoque[self.estoque["Material"] == sap_material]

                console.print(f"\n{material_estoque}", style="italic green")

                if (
                    sap_material not in materiais_receita
                    and sap_material not in self.list_contratada
                    and sap_etapa_material == self.operacao
                ):
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if material_estoque.empty:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )

            if "30002394" not in [i["Material"] for i in material_lista] and not con_met_femea_dn20_estoque.empty:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "ETAPA",
                    self.identificador[1],
                )
                # Adiciona CONEXOES MET LIGACOES FEMEA DN 20.
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "MATERIAL",
                    "30002394",
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "QUANT",
                    "1",
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material,
                    "QUANT",
                )
                ultima_linha_material = ultima_linha_material + 1

            if "30006747" not in [i["Material"] for i in material_lista] and not reg_met_predial_dn20_estoque.empty:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "ETAPA",
                    self.identificador[1],
                )
                # Adiciona REGISTRO METALICO RAMAL PREDIAL DN 20.
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "MATERIAL",
                    "30006747",
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "QUANT",
                    "1",
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material,
                    "QUANT",
                )
                ultima_linha_material = ultima_linha_material + 1

            self.materiais_vigentes()
            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais, self.contrato, self.estoque, self.session)

    def receita_reparo_de_ramal_de_agua(self) -> None:
        """Padrão de materiais no reparo de Ligação de Água."""
        materiais_receita = [
            "30001346",
            "30002394",
            "30001848",
            "300029526",
            "10014709",
            "50000178",
        ]
        sap_material = testa_material_sap.testa_material_sap(self.tb_materiais)
        tubo_pead_dn20_estoque = self.estoque[self.estoque["Material"] == "30001848"]
        if sap_material is None and not tubo_pead_dn20_estoque.empty:
            ultima_linha_material = 0
            self.tb_materiais.InsertRows(str(ultima_linha_material))
            self.tb_materiais.modifyCell(
                ultima_linha_material,
                "ETAPA",
                self.identificador[1],
            )
            # Adiciona TUDO PEAD DN 20.
            self.tb_materiais.modifyCell(
                ultima_linha_material,
                "MATERIAL",
                "30001848",
            )
            self.tb_materiais.modifyCell(
                ultima_linha_material,
                "QUANT",
                "1",
            )
            self.tb_materiais.setCurrentCell(
                ultima_linha_material,
                "QUANT",
            )
            ultima_linha_material = ultima_linha_material + 1
        else:
            material_lista: list[dict[str, str]] = []
            num_material_linhas = self.tb_materiais.RowCount  # Conta as Rows
            # Número da Row do Grid Materiais do SAP
            ultima_linha_material = num_material_linhas
            console.print(f"\nEtapa: {self.operacao}")

            # Loop do Grid Materiais.
            for n_material in range(num_material_linhas):
                sap_material = self.tb_materiais.GetCellValue(n_material, "MATERIAL")
                sap_etapa_material = self.tb_materiais.GetCellValue(n_material, "ETAPA")
                material_lista.append({"Material": sap_material, "Etapa": sap_etapa_material})
                material_estoque = self.estoque[self.estoque["Material"] == sap_material]

                if sap_material not in self.list_contratada:
                    console.print(f"\n{material_estoque}", style="italic green")

                if (
                    sap_material not in materiais_receita
                    and sap_material not in self.list_contratada
                    and sap_etapa_material == self.operacao
                ):
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )
                if material_estoque.empty and sap_material not in self.list_contratada:
                    self.tb_materiais.modifyCheckbox(
                        n_material,
                        "ELIMINADO",
                        True,
                    )

            if "30001848" not in [i["Material"] for i in material_lista] and not tubo_pead_dn20_estoque.empty:
                self.tb_materiais.InsertRows(str(ultima_linha_material))
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "ETAPA",
                    self.identificador[1],
                )
                # Adiciona TUBO PEAD DN 20 - PE 80 - 1.0 MPA.
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "MATERIAL",
                    "30001848",
                )
                self.tb_materiais.modifyCell(
                    ultima_linha_material,
                    "QUANT",
                    "1",
                )
                self.tb_materiais.setCurrentCell(
                    ultima_linha_material,
                    "QUANT",
                )
                ultima_linha_material = ultima_linha_material + 1

            self.materiais_vigentes()
            # Materiais do Global.
            materiais_contratada.materiais_contratada(self.tb_materiais, self.contrato, self.estoque, self.session)
