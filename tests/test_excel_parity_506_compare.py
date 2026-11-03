import importlib.util
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / 'scripts' / 'excel_parity_506_compare.py'
SPEC = importlib.util.spec_from_file_location('excel_parity_506_compare', SCRIPT)
COMPARE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(COMPARE)


class ParityComparatorTests(unittest.TestCase):
    def test_complex_results_compare_by_numeric_components(self):
        gridforge = {'type': 'string', 'value': '1.22464679914735E-16+2i'}
        excel = {'type': 'string', 'value': '1.22514845490862E-16+2i'}
        self.assertTrue(COMPARE.equivalent(gridforge, excel))

    def test_complex_comparison_does_not_hide_material_component_difference(self):
        self.assertFalse(COMPARE.equivalent(
            {'type': 'string', 'value': '1+2i'},
            {'type': 'string', 'value': '1+2.0001i'},
        ))

    def test_complex_comparison_does_not_coerce_non_numeric_strings(self):
        self.assertFalse(COMPARE.equivalent(
            {'type': 'string', 'value': '1+2i'},
            {'type': 'string', 'value': 'not a complex value i'},
        ))


if __name__ == '__main__':
    unittest.main()
