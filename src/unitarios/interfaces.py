# interfaces.py
'''Establish a pattern'''
from abc import ABC, abstractmethod


class UnitarioInterface(ABC):
    '''Interface para as classes'''
    @abstractmethod
    def processar(self):
        '''Processar tse'''

    @abstractmethod
    def pagar(self):
        '''Abstrato para pagar'''
