'''
    Módulo de Testes unitários da aplicação.
'''
# val/src/test/test_main.py

import unittest
from src.core import val


class TestMain(unittest.TestCase):
    '''Class representing for testing'''

    def test_core(self):
        '''Test main core module loop'''
        result = val(["2341816055"], "4600041302", "344")
        correto = "2341816055", 1, True
        self.assertEqual(result, correto, "Resultado não esperado")


if __name__ == '__main__':
    unittest.main()
