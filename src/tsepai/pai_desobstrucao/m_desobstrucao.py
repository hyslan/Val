'''Cabeçalho para serviços de Desobstrução.'''


class Desobstrucao:
    '''Familia DD/DC'''
    MODALIDADE = "desobstrucao"
    OBS = None

    @staticmethod
    def dd_dc():
        '''Serviços de DD/DC, Lavagem, Televisionado'''
        etapa_reposicao = []
        tse_temp_reposicao = []
        tse_proibida = Desobstrucao.OBS
        identificador = Desobstrucao.MODALIDADE
        print("Iniciando processo Pai de DD/DC, Lavagem e Televisionado")

        return tse_temp_reposicao, tse_proibida, identificador, etapa_reposicao
