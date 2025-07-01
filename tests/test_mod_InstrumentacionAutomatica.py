import unittest
from pathlib import Path

class TestModInstrumentacionAutomatica(unittest.TestCase):
    def setUp(self):
        root = Path(__file__).resolve().parents[1]
        bas_path = root / 'Macros' / 'mod_InstrumentacionAutomatica.bas'
        with bas_path.open(encoding='latin-1') as f:
            self.content = f.read()

    def test_contains_function_definitions(self):
        self.assertIn('Sub ExtraerAtributosBloqueInstrumentos', self.content)
        self.assertIn('Sub CompletarSenalesYUnidades', self.content)

    def test_variable_naming(self):
        # Variable should usar el nombre basado en FUNCTION
        self.assertIn('valorFunction', self.content)
        self.assertNotIn('valorFuncion', self.content)

if __name__ == '__main__':
    unittest.main()

