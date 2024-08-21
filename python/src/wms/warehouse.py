"""Módulo onde ficam os materiais disponíveis do dia"""


class Materiais:
    """Classe das bugigangas"""

    ITENS = {
        "abracadeira_reparo": ["30008103", "30002141", "30001120",
                               "30004084", "30002142", "30000255",
                               "30004091", "30001122", "30000153",
                               "30004095", "30002151"],
        "adaptador": ["30008113", "30001136", "30004242",
                      "30002028", "30000518", "30004246"],
        "anel_vedacao_esgoto": ["30002076", "30004141", "30000530"],
        "caixa_ff_tampa_articulada": ["30004388"],
        "caixa_uma": ["50000159", "50000472"],
        "cap": ["30000563", "30004408", "30000032",
                "30004410", "30002355", "30001055",
                "30002356", "30000564"],
        "cavalete_kit": ["30004643"],
        "colar_tomada": ["30002194", "30004691", "30004697",
                         "30000287", "30004701", "30004702",
                         "30004703", "30002202", "30004711",
                         "30001080", "30000288"],
        "colarinho": ["30004732", "30004536", "30004539",
                      "30002222", "30000249", "30004556"],
        "conexoes_metalicas": ["30001346", "30002393", "30002394"],
        "cotovelo_45": ["30004940"],
        "cotovelo_90": ["30004944", "30004945", "30000001"],
        "curva_45": ["30005264", "30005265", "30008307"],
        "curva_45_esgoto": ["30005282", "30005285", "30005282",
                            "30022734"],
        "curva_90": ["30002714", "30000051", "30005330",
                     "30008324"],
        "curva_90_esgoto": ["30008322", "30005146", "30005147",
                            "30005148", "30002722"],
        "dispositivo_medicao": ["50000021", "50000178"],
        "extremidade": ["30005509", "30005510", "30001252",
                        "30005514", "30005517", "30000777",
                        "30002609", "30005519"],
        "guarnicao": ["50000202", "50000098"],
        "hidrante": ["30002899"],
        "hidrometro": ["50000108", "50000530", "50000221",
                       "50000387", "50000057", "50000113",
                       "50000252"],
        "junta_agua": ["30005601", "30005602", "30002952",
                       "30005611", "30002956", "30000614"],
        "junta_esgoto": ["30005615", "30001528", "30002958",
                         "30005617", "30000357", "30005619"
                         "30001529"],
        "lacre": ["50001070"],
        "luva_ajustavel": ["30005876"],
        "luva_bipartida": ["30005887", "30001546", "30005888",
                           "30002794", "30005893"],
        "luva_correr_esgoto": ["30005898", "30008427", "30002798",
                               "30005896", "30001548", "30002797",
                               "30005894"],
        "luva_correr_bolsa": ["30005906", "30002805", "30005917",
                              "30002809", "30005919", "30001554",
                              "30005920", "30005924", "30002812",
                              "30008432", "30002813", "30008432",
                              "30002813"],
        "luva_correr_agua": ["30008433", "30001557"],
        "luva_eletrofusao": ["30005746", "30002823", "30005747",
                             "30001561", "30005748", "30005750"],
        "luva_ff": ["30005771", "30005772", "30005773"],
        "luva_reducao_ff": ["30002845", "30005791"],
        "porca": ["50000139", "50000285"],
        "reducao": ["30001585", "30006864", "30008609",
                    "30010436", "30006911", "30000206",
                    "30006933"],
        "registro": ["30003522", "30021805", "30006747"],
        "reparador_asfaltico": ["10014709"],
        "selim": ["30007132", "30003416", "30023927"],
        "suporte_adaptador_plastico": ["50000143"],
        "tampao_agua": ["30003442", "30003439", "30006984",
                        "30006983", "30032251", "30032233"],
        "tampao_esgoto": ["30032220"],
        "te_agua": ["30008671", "30000211", "30008677",
                    "30007034", "30003467", "30007235",
                    "30001683", "30000992", "30003669",
                    "30007245", "30003672", "30008683",
                    "30001891", "30003715", "30007286",
                    "30001890", "30003693", "30000804",
                    "30001914", "30003730", "30000213",
                    "30007195", "30007229", "30007233",
                    "30008708"],
        "te_esgoto": ["30007203", "30001925", "30007206",
                      "30003756"],
        "tubete": ["50000037", "50000305"],
        "tubo_agua": ["30003896", "30000849", "30007697",
                      "30000125", "30000851", "30003905",
                      "30007626", "30007628", "30028862",
                      "30007888", "30008808", "30003797",
                      "30001848", "30007896", "30008810",
                      "30007909", "30026318"],
        "tubo_esgoto": ["30003817", "30007933", "30028892",
                        "30028858", "30028857", "30028856",
                        "30007933", "30003817"],
        "uniao": ["30003832", "30007765"],
        "valvula": ["30001875", "30007804", "30003852",
                    "30007805", "30003853", "30007807",
                    "30001876", "30007808", "30003854",
                    "30007809", "30007810", "30001877",
                    "30007812", "30003856", "30003860"],

    }

    def __init__(self, estoque) -> None:
        self.estoque = estoque

    def _verificar_disponibilidade(self, itens):
        """Verifica se tem disponível o material no estoque."""
        return [item for item in itens if item in self.estoque]

    def obter_itens_disponiveis(self, tipo_item):
        """Retorna os itens disponíveis no estoque"""
        itens = self.ITENS.get(tipo_item, [])
        return self._verificar_disponibilidade(itens)

    def reparador_asfaltico(self):
        """Asfalto Frio"""
        return self.obter_itens_disponiveis("reparador_asfaltico")


class Agua(Materiais):
    """Família Água"""

    def __init__(self, estoque) -> None:
        super().__init__(estoque)

    def abracadeira_reparo(self):
        """Abraçadeira de Ferro para Reparo"""
        return self.obter_itens_disponiveis("abracadeira_reparo")

    def adaptador(self):
        """Adaptador para PVC/FF"""
        return self.obter_itens_disponiveis("adaptador")

    def caixa_ff_tampa_articulada(self):
        """Caixa Tampa Articulada"""
        return self.obter_itens_disponiveis("caixa_ff_tampa_articulada")

    def caixa_uma(self):
        """Caixa UMA - Cavalete"""
        return self.obter_itens_disponiveis("caixa_uma")

    def cap(self):
        """CAP - Rede Água"""
        return self.obter_itens_disponiveis("cap")

    def cavalete_kit(self):
        """Kit Cavalete Completo"""
        return self.obter_itens_disponiveis("cavalete_kit")

    def colar_tomada(self):
        """Colar Tomada"""
        return self.obter_itens_disponiveis("colar_tomada")

    def colarinho(self):
        """Colarinho"""
        return self.obter_itens_disponiveis("colarinho")

    def cotovelo_90(self):
        """Cotovelo 90 graus"""
        return self.obter_itens_disponiveis("cotovelo_90")

    def conexoes_metalicas(self):
        """Conexões MET Macho/Fêmea"""
        return self.obter_itens_disponiveis("conexoes_metalicas")

    def cotovelo_45(self):
        """Cotovelo 45 graus"""
        return self.obter_itens_disponiveis("cotovelo_45")

    def curva_45(self):
        """Curva 45 graus água"""
        return self.obter_itens_disponiveis("curva_45")

    def curva_90(self):
        """Curva 90 graus água"""
        return self.obter_itens_disponiveis("curva_90")

    def dispositivo_medicao(self):
        """Dispositivo de Medição - água"""
        return self.obter_itens_disponiveis("dispositivo_medicao")

    def extremidade(self):
        """Extremidade - água"""
        return self.obter_itens_disponiveis("extremidade")

    def guarnicao(self):
        """Guarnição água"""
        return self.obter_itens_disponiveis("guarnicao")

    def hidrante(self):
        """Hidrante"""
        return self.obter_itens_disponiveis("hidrante")

    def hidrometro(self):
        """Hidrômetros"""
        return self.obter_itens_disponiveis("hidrometro")

    def junta_agua(self):
        """Junta para água"""
        return self.obter_itens_disponiveis("junta_agua")

    def lacre(self):
        """Lacre do cavalete."""
        return self.obter_itens_disponiveis("lacre")

    def luva_ajustavel(self):
        """Luva Ajustável"""
        return self.obter_itens_disponiveis("luva_ajustavel")

    def luva_bipartida(self):
        """Luva Bipartida"""
        return self.obter_itens_disponiveis("luva_bipartida")

    def luva_correr_bolsa(self):
        """Luva de Correr com Bolsa"""
        return self.obter_itens_disponiveis("luva_correr_bolsa")

    def luva_correr_agua(self):
        """Luva de Correr de Água PVC"""
        return self.obter_itens_disponiveis("luva_correr_agua")

    def luva_eletrofusao(self):
        """Luva de Eletrofusão - Água"""
        return self.obter_itens_disponiveis("luva_eletrofusao")

    def luva_ff(self):
        """Luva de ferro"""
        return self.obter_itens_disponiveis("luva_ff")

    def luva_reducao_ff(self):
        """Luva com redução de ferro"""
        return self.obter_itens_disponiveis("luva_reducao_ff")

    def porca(self):
        """Porca - água"""
        return self.obter_itens_disponiveis("porca")

    def reducao(self):
        """Redução"""
        return self.obter_itens_disponiveis("reducao")

    def registro(self):
        """Registro"""
        return self.obter_itens_disponiveis("registro")

    def suporte_adaptador_plastico(self):
        """Suporte Adaptador Plástico p/ dispositivo"""
        return self.obter_itens_disponiveis("suporte_adaptador_plastico")

    def tampao_agua(self):
        """Tampão articulado - água"""
        return self.obter_itens_disponiveis("tampao_agua")

    def te_agua(self):
        """TE - água"""
        return self.obter_itens_disponiveis("te_agua")

    def tubete(self):
        """Tubete"""
        return self.obter_itens_disponiveis("tubete")

    def tubo_agua(self):
        """Tubo PEAD"""
        return self.obter_itens_disponiveis("tubo_agua")

    def uniao(self):
        """União para PEAD"""
        return self.obter_itens_disponiveis("uniao")

    def valvula(self):
        """Válvula GAV"""
        return self.obter_itens_disponiveis("valvula")


class Esgoto(Materiais):
    """Família Esgoto"""

    def __init__(self, estoque) -> None:
        super().__init__(estoque)

    def anel_vedacao_esgoto(self):
        """Anel de Vedação PVC/FF"""
        return self.obter_itens_disponiveis("anel_vedacao_esgoto")

    def curva_45_esgoto(self):
        """Curva 45 graus esgoto"""
        return self.obter_itens_disponiveis("curva_45_esgoto")

    def curva_90_esgoto(self):
        """Curva 90 graus esgoto"""
        return self.obter_itens_disponiveis("curva_90_esgoto")

    def junta_esgoto(self):
        """Junta para esgoto"""
        return self.obter_itens_disponiveis("junta_esgoto")

    def luva_correr_esgoto(self):
        """Luva de Correr de Esgoto PVC"""
        return self.obter_itens_disponiveis("luva_correr_esgoto")

    def selim(self):
        """Selim - Esgoto"""
        return self.obter_itens_disponiveis("selim")

    def tampao_esgoto(self):
        """Tampão articulado - esgoto"""
        return self.obter_itens_disponiveis("tampao_esgoto")

    def te_esgoto(self):
        """TE - Esgoto"""
        return self.obter_itens_disponiveis("te_esgoto")

    def tubo_esgoto(self):
        """Tubos PVC de Esgoto"""
        return self.obter_itens_disponiveis("tubo_esgoto")
