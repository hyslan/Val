from wms.consulta_estoque import estoque_novasp
import sap

sessions = sap.listar_sessoes()
session = sap.criar_sessao(sessions)
estoque = estoque_novasp(session, sessions)

