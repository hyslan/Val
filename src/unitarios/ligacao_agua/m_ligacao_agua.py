'''Módulo Família Ligação Água Unitário.'''
from src.unitarios.base import BaseUnitario
from src.unitarios.localizador import btn_localizador
from src.lista_reposicao import dict_reposicao


class LigacaoAgua(BaseUnitario):
    '''Ramo de Ligações (Ramal) de água'''

    def reposicoes(self):
        '''Reposições dos serviços de Ligação de água'''
        pass

    def ligacao_agua(self):
        '''Ramal novo de água, avulsa.'''
        pass

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
